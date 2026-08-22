from __future__ import annotations

import csv
import hashlib
import json
import shutil
import sys
import unicodedata
import uuid
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

from .configuration import ConfigurationError, load_config, resolve_from_config
from .paths import project_root, resource_root
from .sod_analysis import analyze as analyze_sod, export_analysis
from .template_exporter import export_template


LOCAL_XLRD = resource_root() / ".build_tools" / "xlrd_runtime"
if LOCAL_XLRD.is_dir() and str(LOCAL_XLRD) not in sys.path:
    sys.path.insert(0, str(LOCAL_XLRD))


class SourceValidationError(ValueError):
    """Indica uma estrutura de fonte que não pode ser mapeada com segurança."""


@dataclass(frozen=True)
class ProcessingResult:
    run_id: str
    workbook_path: Path | None
    log_path: Path
    pending_path: Path
    validation_path: Path
    diagnostics_path: Path
    email_draft_path: Path
    sod_analysis_path: Path | None
    system_mapping_path: Path
    user_conflicts_path: Path
    email_required: bool
    status: str


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _key(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _header_key(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value)).encode("ascii", "ignore").decode("ascii")
    return " ".join(text.casefold().split())


def _system_key(value) -> str:
    return _header_key(_key(value)).replace(" - ", " ").replace("_", " ")


def _engine(path: Path) -> str | None:
    return "xlrd" if path.suffix.casefold() == ".xls" else None


def _read_excel(path: Path, sheet: str | None = None):
    try:
        with pd.ExcelFile(path, engine=_engine(path)) as book:
            target = sheet or book.sheet_names[0]
            if target not in book.sheet_names:
                raise ConfigurationError(f"Aba '{target}' não encontrada em {path.name}.")
            return target, pd.read_excel(book, sheet_name=target, dtype=object)
    except ImportError as exc:
        if path.suffix.casefold() == ".xls":
            raise SourceValidationError(f"A base legado '{path.name}' exige o leitor xlrd. Instale as dependências do projeto antes de processar arquivos .xls.") from exc
        raise


def _resolve_columns(frame: pd.DataFrame, spec: dict, label: str) -> dict[str, str]:
    available = {str(column): column for column in frame.columns}
    normalised: dict[str, list[str]] = {}
    for column in available:
        normalised.setdefault(_header_key(column), []).append(column)
    selected: dict[str, str] = {}
    unresolved: list[str] = []
    for canonical, source in spec.get("colunas", {}).items():
        alternatives = [source] if isinstance(source, str) else list(source)
        exact = next((str(item) for item in alternatives if str(item) in available), None)
        if exact:
            selected[canonical] = exact
            continue
        matches = list(dict.fromkeys(match for item in alternatives for match in normalised.get(_header_key(item), [])))
        if len(matches) == 1:
            selected[canonical] = matches[0]
        elif canonical in spec.get("campos_obrigatorios", []):
            unresolved.append(canonical)
    if unresolved:
        headers = ", ".join(str(column) for column in frame.columns)
        raise SourceValidationError(f"Não foi possível localizar os campos obrigatórios em {label}: {', '.join(unresolved)}. Colunas encontradas: {headers}. Ajuste o JSON de mapeamento, sem alterar a fonte.")
    return selected


def _load_source(path: Path, spec: dict, label: str) -> tuple[list[dict], dict]:
    if not path.is_file():
        raise FileNotFoundError(f"Arquivo de {label} não encontrado: {path}")
    sheet, frame = _read_excel(path, spec.get("aba"))
    selected = _resolve_columns(frame, spec, label)
    records: list[dict] = []
    for position, row in frame.iterrows():
        record = {canonical: (_key(row.get(source)) or None) if canonical in selected else None for canonical, source in selected.items()}
        for canonical in spec.get("colunas", {}):
            record.setdefault(canonical, None)
        record["_linha_origem"] = int(position) + 2
        record["_arquivo_origem"] = path.name
        records.append(record)
    required = spec.get("campos_obrigatorios", [])
    blank_required = sum(1 for row in records for field in required if not _key(row.get(field)))
    return records, {"arquivo": str(path), "aba": sheet, "linhas_lidas": len(records), "duplicidades_exatas": int(frame.duplicated().sum()), "campos_obrigatorios_vazios": blank_required, "mapeamento_aplicado": selected, "sha256": _sha256(path)}


def _source_blank_issues(records: list[dict], required: list[str], source_label: str) -> list[dict]:
    """Registra cada campo indispensável vazio sem esconder a linha de origem."""
    issues: list[dict] = []
    for row in records:
        for field in required:
            if _key(row.get(field)):
                continue
            is_functions = source_label == "Base de funcionalidades"
            issues.append({
                "tipo": "CAMPO_OBRIGATORIO_VAZIO",
                "sistema_usuarios": "" if is_functions else _key(row.get("sistema")),
                "sistema_funcionalidades": _key(row.get("sistema")) if is_functions else "",
                "perfil": _key(row.get("perfil")),
                "onde": f"{source_label}: {row['_arquivo_origem']}, linha {row['_linha_origem']}, campo {field}",
                "como_corrigir": f"Preencha o campo obrigatório '{field}' na origem ou ajuste o alias correspondente no JSON de mapeamento; a fonte RAW da execução foi preservada.",
            })
    return issues


def _group(records: list[dict], field: str) -> dict[str, list[dict]]:
    groups: dict[str, list[dict]] = {}
    for record in records:
        groups.setdefault(_system_key(record.get(field)) or "__SEM_SISTEMA__", []).append(record)
    return groups


def _profiles(records: list[dict]) -> set[str]:
    return {_key(row.get("perfil")) for row in records if _key(row.get("perfil"))}


def _system_label(groups: dict[str, list[dict]], key: str) -> str:
    if key == "__SEM_SISTEMA__":
        return "(não informado)"
    return _key(groups[key][0].get("sistema")) or key


def _target_from_user_filename(user_rows: list[dict], function_groups: dict[str, list[dict]]) -> str | None:
    """Usa o identificador explícito do arquivo somente quando ele aponta um único sistema."""
    names = {_system_key(Path(_key(row.get("_arquivo_origem"))).stem) for row in user_rows if _key(row.get("_arquivo_origem"))}
    candidates: set[str] = set()
    for name in names:
        compact = " ".join(token for token in name.split() if token not in {"base", "usuario", "usuarios", "perfil", "perfis", "xls", "xlsx"})
        if not compact:
            continue
        candidates.update(key for key in function_groups if key == compact or key.startswith(compact + " "))
    return next(iter(candidates)) if len(candidates) == 1 else None


def _target_from_user_entitlements(user_rows: list[dict], function_groups: dict[str, list[dict]]) -> str | None:
    """Infere o sistema pelo prefixo de entitlement apenas quando for inequívoco.

    Exemplo: direitos `CAC_...` podem apontar para `CAC - CENTRAL
    DEPOSITARIA`. A inferência não é aplicada se houver mais de um candidato.
    """
    prefixes = {
        _system_key(row.get("entitlement")).split(" ")[0]
        for row in user_rows
        if _system_key(row.get("entitlement"))
    }
    prefixes.discard("")
    candidates = {
        system
        for prefix in prefixes
        for system in function_groups
        if system == prefix or system.startswith(prefix + " ")
    }
    return next(iter(candidates)) if len(candidates) == 1 else None


def _map_systems(functions: list[dict], users: list[dict], config: dict) -> tuple[list[dict], list[dict], list[dict]]:
    """Relaciona sistemas por nome, de-para aprovado ou interseção única de perfis."""
    function_groups, user_groups = _group(functions, "sistema"), _group(users, "sistema")
    approved = {_system_key(source): _system_key(target) for source, target in config.get("de_para_sistemas", {}).items()}
    mappings: list[dict] = []
    issues: list[dict] = []
    contexts: list[dict] = []
    used_function_groups: set[str] = set()
    for user_system, user_rows in user_groups.items():
        target = user_system if user_system in function_groups else approved.get(user_system)
        method = "nome_normalizado" if target == user_system and target in function_groups else "de_para_configurado"
        if not target or target not in function_groups:
            target_from_file = _target_from_user_filename(user_rows, function_groups) if user_system == "__SEM_SISTEMA__" else None
            if target_from_file:
                target, method = target_from_file, "arquivo_origem"
        if not target or target not in function_groups:
            target_from_entitlement = _target_from_user_entitlements(user_rows, function_groups) if user_system == "__SEM_SISTEMA__" else None
            if target_from_entitlement:
                target, method = target_from_entitlement, "prefixo_entitlement_unico"
        if not target or target not in function_groups:
            overlap = [name for name, rows in function_groups.items() if _profiles(user_rows) & _profiles(rows)]
            if len(overlap) == 1:
                target, method = overlap[0], "perfil_em_comum_unico"
            elif len(overlap) > 1:
                issues.append({"tipo": "SISTEMA_AMBIGUO", "sistema_usuarios": _system_label(user_groups, user_system), "sistema_funcionalidades": "", "perfil": "", "onde": "De-para de sistemas", "como_corrigir": "Informe o de-para de sistemas no JSON; mais de um sistema de funcionalidades contém perfis da base de usuários."})
                continue
            else:
                issues.append({"tipo": "SISTEMA_SEM_CORRESPONDENCIA", "sistema_usuarios": _system_label(user_groups, user_system), "sistema_funcionalidades": "", "perfil": "", "onde": "De-para de sistemas", "como_corrigir": "Envie a base de funcionalidades correspondente ou cadastre o de-para de sistemas aprovado."})
                continue
        used_function_groups.add(target)
        mappings.append({"sistema_usuarios": _system_label(user_groups, user_system), "sistema_funcionalidades": _system_label(function_groups, target), "metodo": method, "perfis_usuarios": len(_profiles(user_rows)), "perfis_funcionalidades": len(_profiles(function_groups[target]))})
        user_profiles, function_profiles = _profiles(user_rows), _profiles(function_groups[target])
        for profile in sorted(user_profiles - function_profiles):
            issues.append({"tipo": "PERFIL_USUARIO_SEM_FUNCIONALIDADES", "sistema_usuarios": _system_label(user_groups, user_system), "sistema_funcionalidades": _system_label(function_groups, target), "perfil": profile, "onde": "Base de usuários", "como_corrigir": "Confirmar se o perfil está ativo e enviar as transações/funcionalidades correspondentes, ou formalizar que está fora de escopo."})
        for profile in sorted(function_profiles - user_profiles):
            issues.append({"tipo": "PERFIL_FUNCIONALIDADE_SEM_USUARIOS", "sistema_usuarios": _system_label(user_groups, user_system), "sistema_funcionalidades": _system_label(function_groups, target), "perfil": profile, "onde": "Base de funcionalidades", "como_corrigir": "Confirmar se o perfil possui usuários ativos no escopo ou formalizar que está inativo, obsoleto ou fora de escopo."})
        contexts.append({"sistema_funcionalidades": target, "sistema_usuarios": user_system, "usuarios": user_rows})
    for function_system in sorted(set(function_groups) - used_function_groups):
        issues.append({"tipo": "SISTEMA_FUNCIONALIDADE_SEM_BASE_USUARIOS", "sistema_usuarios": "", "sistema_funcionalidades": _system_label(function_groups, function_system), "perfil": "", "onde": "Base de funcionalidades", "como_corrigir": "Envie a base de usuários do sistema ou formalize que não há usuários no escopo."})
    return mappings, issues, contexts


def _matrix_rows(functions: list[dict]) -> list[dict]:
    rows: list[dict] = []
    for number, source in enumerate(functions, start=1):
        row = dict(source)
        row["id"] = f"FUNC_{number:04d}"
        row["id_origem"] = row["id"]
        # A descrição da atividade é a evidência operacional da fonte. Só
        # construímos um texto de apoio quando ela realmente não foi fornecida;
        # assim LINE, CAC ou qualquer outro sistema não recebem nomes genéricos.
        activity = _key(row.get("atividade")) or _key(row.get("funcionalidade"))
        transaction = _key(row.get("transacao_tela_direito"))
        if activity:
            row["atividade"] = activity
        elif transaction:
            row["atividade"] = f"Executar a funcionalidade {transaction}"
        for field in ("area", "responsabilidade_area", "gestor_area", "email_gestor", "modulo"):
            if not _key(row.get(field)):
                row[field] = "-"
        rows.append(row)
    return rows


def _observed_conflicts(matrix: list[dict], contexts: list[dict], rules: list[dict], config: dict) -> tuple[list[dict], dict, list[dict]]:
    activity_field = config.get("confronto", {}).get("campo_id_atividade", "id")
    first_field, second_field = "id_atividade_1", "id_atividade_2"
    activities_by_system_profile: dict[tuple[str, str], set[str]] = {}
    labels: dict[str, set[str]] = {}
    for row in matrix:
        system, profile, identifier = _system_key(row.get("sistema")), _key(row.get("perfil")), _key(row.get(activity_field))
        if system and profile and identifier:
            activities_by_system_profile.setdefault((system, profile), set()).add(identifier)
            labels.setdefault(identifier, set()).update(filter(None, (_key(row.get("atividade")), _key(row.get("funcionalidade")), _key(row.get("transacao_tela_direito")))))
    results: list[dict] = []
    evidence: list[dict] = []
    seen: set[tuple[str, str]] = set()
    symmetric = divergent = 0
    for rule in rules:
        first, second = _key(rule.get(first_field)), _key(rule.get(second_field))
        if not first or not second:
            continue
        rule_label_1, rule_label_2 = _key(rule.get("atividade_1")), _key(rule.get("atividade_2"))
        if (rule_label_1 and labels.get(first) and rule_label_1 not in labels[first]) or (rule_label_2 and labels.get(second) and rule_label_2 not in labels[second]):
            divergent += 1
            continue
        affected: list[tuple[str, str, list[str]]] = []
        for context in contexts:
            system = context["sistema_funcionalidades"]
            for user_row in context["usuarios"]:
                user, profile = _key(user_row.get("usuario")), _key(user_row.get("perfil"))
                activities = activities_by_system_profile.get((system, profile), set())
                if user and first in activities and second in activities:
                    affected.append((user, system, [profile]))
        if not affected:
            continue
        signature = tuple(sorted((first, second)))
        if signature in seen:
            symmetric += 1
            continue
        seen.add(signature)
        results.append(dict(rule))
        for user, system, profiles in affected:
            evidence.append({"usuario": user, "sistema": system, "perfis": " | ".join(profiles), "id_conflito": _key(rule.get("id_conflito")), "id_atividade_1": first, "atividade_1": rule_label_1 or next(iter(labels.get(first, {""}))), "id_atividade_2": second, "atividade_2": rule_label_2 or next(iter(labels.get(second, {""})))})
    return results, {"regras_aplicadas": len(results), "duplicidades_simetrica_sinalizadas": symmetric, "regras_com_atividade_divergente": divergent, "usuarios_afetados": len(evidence)}, evidence


def _sat_rows(matrix: list[dict], config: dict, pending: list[str]) -> list[dict]:
    sat = config.get("sat", {})
    if not sat.get("habilitado", False):
        pending.append("SAT não preenchido: o critério aprovado está desabilitado na configuração.")
        return []
    field, values = sat.get("campo_origem"), {str(item) for item in sat.get("valores_aprovados", [])}
    if not field or not values:
        pending.append("SAT não preenchido: campo ou valores aprovados do critério estão ausentes.")
        return []
    output = []
    for row in matrix:
        if _key(row.get(field)) in values:
            item = dict(row)
            item.setdefault("descricao_atividade", row.get("atividade"))
            output.append(item)
    return output


def _write_csv(path: Path, columns: list[str], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def _write_validation(
    path: Path,
    mappings: list[dict],
    issues: list[dict],
    blocked: bool = False,
    blocked_by_profile_divergence: bool = False,
) -> None:
    lines = ["# Teste de consistência das bases", "", "A validação foi executada por sistema e nos dois sentidos: usuários → funcionalidades e funcionalidades → usuários.", "", f"Sistemas conciliados: {len(mappings)}", f"Pendências identificadas: {len(issues)}", "", "## De-para aplicado", ""]
    lines.extend(f"- `{row['sistema_usuarios']}` → `{row['sistema_funcionalidades']}` ({row['metodo']})" for row in mappings)
    lines.extend(["", "## Resultado", ""])
    if blocked:
        lines.append("- Processamento bloqueado: não houve conciliação segura entre as bases. Nenhum Excel de resultado foi gerado para evitar uma matriz vazia ou de sistema incorreto.")
    elif blocked_by_profile_divergence:
        lines.append("- Matriz Funcional final não gerada: foram comprovadas divergências de perfis entre as bases conciliadas. Consulte o diagnóstico, a minuta de e-mail e a análise SoD preliminar antes de uma nova execução.")
    elif issues:
        lines.append("- A geração da Matriz Funcional foi mantida. Revise `diagnostico_inconsistencias.csv` e `minuta_email_inconsistencias.md` antes da validação final de Conflitos e SAT.")
    else:
        lines.append("- Não foram identificadas inconsistências de sistema ou perfil.")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _write_email_draft(path: Path, issues: list[dict]) -> None:
    """Gera uma minuta específica apenas quando a conciliação comprovou perfis divergentes."""
    grouped: dict[str, dict[str, list[str]]] = {}
    for item in issues:
        if item["tipo"] not in {"PERFIL_USUARIO_SEM_FUNCIONALIDADES", "PERFIL_FUNCIONALIDADE_SEM_USUARIOS"}:
            continue
        system = item["sistema_funcionalidades"] or item["sistema_usuarios"]
        group = grouped.setdefault(system, {"usuarios_sem_funcionalidades": [], "funcionalidades_sem_usuarios": []})
        key = "usuarios_sem_funcionalidades" if item["tipo"] == "PERFIL_USUARIO_SEM_FUNCIONALIDADES" else "funcionalidades_sem_usuarios"
        group[key].append(item["perfil"])
    systems = sorted(grouped)
    subject_system = systems[0] if len(systems) == 1 else "Sistemas analisados"
    lines = ["# Minuta de e-mail — divergência comprovada de perfis", "", f"Assunto: [Projeto B3] - SoD - Sistema {subject_system}", "", "Olá, pessoal. Boa tarde!", "", f"Durante a análise das informações do sistema {subject_system}, confrontamos os perfis associados aos usuários na base recebida com os perfis da matriz de transações, nos dois sentidos. Foram identificadas as divergências abaixo:", ""]
    if not grouped:
        lines.extend(["Não foram encontradas divergências de perfis conciliados que demandem comunicação externa.", ""])
    for system in systems:
        profile_sets = grouped[system]
        lines.extend([f"## Sistema: {system}", ""])
        if profile_sets["usuarios_sem_funcionalidades"]:
            lines.extend(["Perfis associados aos usuários que não foram localizados na matriz de transações:", ""])
            lines.extend(f"- {profile}" for profile in sorted(set(profile_sets["usuarios_sem_funcionalidades"])))
            lines.append("")
        if profile_sets["funcionalidades_sem_usuarios"]:
            lines.extend(["Perfis existentes na matriz de transações que não foram localizados na base de usuários:", ""])
            lines.extend(f"- {profile}" for profile in sorted(set(profile_sets["funcionalidades_sem_usuarios"])))
            lines.append("")
    if grouped:
        lines.extend(["Poderiam, por gentileza, confirmar se os perfis relacionados permanecem ativos e estão no escopo da análise? Quando aplicável, solicitamos o envio da relação de transações/funcionalidades ou da base de usuários correspondente. Para perfis obsoletos, inativos ou fora de escopo, solicitamos a formalização dessa condição.", "", "Essa confirmação é necessária para concluir a Matriz Funcional e as análises de Segregação de Funções (SoD).", "", "Ficamos no aguardo.", "", "Obrigada."])
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def _pending_text(run_id: str, pending: list[str]) -> str:
    return "\n".join(["# Pendências de processamento", "", f"Execução: `{run_id}`", "", *(f"- {item}" for item in (pending or ["Nenhuma pendência registrada."]))]) + "\n"


def _safe_output_name(configured_name: str, systems: set[str]) -> str:
    """Inclui o sistema conciliado no nome sem usar caracteres inválidos no Windows."""
    suffix = next(iter(systems)) if len(systems) == 1 else "Multissistemas" if systems else "Sistema_nao_identificado"
    safe = "".join("_" if character in '<>:"/\\|?*' else character for character in str(suffix)).strip().rstrip(".")
    safe = "_".join(safe.split()) or "Sistema_nao_identificado"
    base = Path(configured_name).stem or "Matriz_SoD_SAT_resultado"
    return f"{base}_{safe}.xlsx"


def process_workbooks(funcionalidades_path: Path, usuarios_path: Path, conflitos_path: Path | None, config_path: Path, output_root: Path) -> ProcessingResult:
    config = load_config(config_path)
    funcionalidades_path, usuarios_path = funcionalidades_path.resolve(), usuarios_path.resolve()
    conflitos_path = conflitos_path.resolve() if conflitos_path else None
    if funcionalidades_path == usuarios_path or (conflitos_path and conflitos_path in {funcionalidades_path, usuarios_path}):
        raise ValueError("Selecione fontes distintas para funcionalidades, usuários e regras SoD.")
    run_id = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ") + "_" + uuid.uuid4().hex[:8]
    run_dir, raw_dir = output_root / "execucoes" / run_id, output_root / "execucoes" / run_id / "raw"
    raw_dir.mkdir(parents=True, exist_ok=False)
    functions, functions_summary = _load_source(funcionalidades_path, config["base_funcionalidades"], "base de funcionalidades")
    users, users_summary = _load_source(usuarios_path, config["base_usuarios"], "base de usuários")
    shutil.copy2(funcionalidades_path, raw_dir / funcionalidades_path.name)
    shutil.copy2(usuarios_path, raw_dir / usuarios_path.name)
    if conflitos_path:
        rules, rules_summary = _load_source(conflitos_path, config["atividades_conflitantes"], "regras de conflitos")
        shutil.copy2(conflitos_path, raw_dir / conflitos_path.name)
    else:
        rules, rules_summary = [], {"arquivo": None, "linhas_lidas": 0, "nao_fornecida": True}
    mappings, reconciliation_issues, contexts = _map_systems(functions, users, config)
    source_issues = _source_blank_issues(functions, config["base_funcionalidades"].get("campos_obrigatorios", []), "Base de funcionalidades")
    source_issues.extend(_source_blank_issues(users, config["base_usuarios"].get("campos_obrigatorios", []), "Base de usuários"))
    issues = reconciliation_issues + source_issues
    systems_in_scope = {context["sistema_funcionalidades"] for context in contexts}
    raw_user_systems = {_key(row.get("sistema")) for row in users if _key(row.get("sistema"))}
    raw_function_systems = {_key(row.get("sistema")) for row in functions if _key(row.get("sistema"))}
    raw_scoped_systems = {_key(row.get("sistema")) for row in functions if _system_key(row.get("sistema")) in systems_in_scope}
    # Quando não existe de-para seguro, não se pode preencher uma matriz LINE
    # com funcionalidades de CAC/STAR/etc. A saída fica vazia e a divergência
    # é documentada para tratativa, sem criar um falso positivo editorial.
    functions_in_scope = [row for row in functions if _system_key(row.get("sistema")) in systems_in_scope] if systems_in_scope else []
    display_systems = raw_scoped_systems or raw_user_systems or raw_function_systems
    display_label = next(iter(display_systems)) if len(display_systems) == 1 else "Múltiplos sistemas" if display_systems else "Sistema não identificado"
    profile_divergence_issues = [
        item for item in reconciliation_issues
        if item["tipo"] in {"PERFIL_USUARIO_SEM_FUNCIONALIDADES", "PERFIL_FUNCIONALIDADE_SEM_USUARIOS"}
    ]
    validation_path, diagnostics_path = run_dir / "teste_consistencia_bases.md", run_dir / "diagnostico_inconsistencias.csv"
    email_draft_path, system_mapping_path = run_dir / "minuta_email_inconsistencias.md", run_dir / "de_para_sistemas.csv"
    user_conflicts_path = run_dir / "evidencias_conflitos_por_usuario.csv"
    log_path, pending_path = run_dir / "log_processamento.json", run_dir / "pendencias.md"
    if not contexts or not functions_in_scope:
        pending = [
            "Processamento bloqueado: não foi possível conciliar ao menos um sistema entre as bases com evidência suficiente.",
            "Nenhum Excel de resultado foi gerado para evitar uma matriz vazia, genérica ou preenchida com dados de outro sistema.",
            "Consulte o diagnóstico para revisar o de-para de sistemas, os perfis e a estrutura das fontes.",
        ]
        _write_validation(validation_path, mappings, issues, blocked=True)
        _write_csv(diagnostics_path, ["tipo", "sistema_usuarios", "sistema_funcionalidades", "perfil", "onde", "como_corrigir"], issues)
        _write_email_draft(email_draft_path, issues)
        _write_csv(system_mapping_path, ["sistema_usuarios", "sistema_funcionalidades", "metodo", "perfis_usuarios", "perfis_funcionalidades"], mappings)
        _write_csv(user_conflicts_path, ["usuario", "sistema", "perfis", "id_conflito", "id_atividade_1", "atividade_1", "id_atividade_2", "atividade_2"], [])
        log = {
            "run_id": run_id,
            "gerado_em_utc": datetime.now(timezone.utc).isoformat(),
            "status": "BLOQUEADO_SEM_CONCILIACAO",
            "template_oficial": {"arquivo": str(resolve_from_config(config_path, config["template_oficial"]["arquivo"])), "sha256": _sha256(resolve_from_config(config_path, config["template_oficial"]["arquivo"]))},
            "base_funcionalidades": functions_summary,
            "base_usuarios": users_summary,
            "de_para_sistemas": mappings,
            "escopo_matriz_funcional": {"sistema_exibido": display_label, "registros_entrada": len(functions), "registros_no_escopo": 0},
            "inconsistencias": {"quantidade": len(issues), "sistema_perfil": len(reconciliation_issues), "qualidade_fonte": len(source_issues), "arquivo": str(diagnostics_path)},
            "pendencias": pending,
        }
        log_path.write_text(json.dumps(log, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")
        pending_path.write_text(_pending_text(run_id, pending), encoding="utf-8")
        return ProcessingResult(run_id, None, log_path, pending_path, validation_path, diagnostics_path, email_draft_path, None, system_mapping_path, user_conflicts_path, False, "BLOQUEADO_SEM_CONCILIACAO")
    matrix = _matrix_rows(functions_in_scope)
    pending: list[str] = []
    if functions_summary["campos_obrigatorios_vazios"]:
        pending.append(f"Base de funcionalidades contém {functions_summary['campos_obrigatorios_vazios']} valor(es) vazio(s) em campos obrigatórios.")
    if users_summary["campos_obrigatorios_vazios"]:
        pending.append(f"Base de usuários contém {users_summary['campos_obrigatorios_vazios']} valor(es) vazio(s) em campos obrigatórios.")
    if reconciliation_issues:
        pending.append(f"Foram identificadas {len(reconciliation_issues)} pendência(s) de sistema/perfil. Consulte o diagnóstico e a minuta de e-mail.")
    if source_issues:
        pending.append(f"Foram identificadas {len(source_issues)} pendência(s) de qualidade da fonte. Consulte o diagnóstico para arquivo, linha, campo e correção necessária.")
    if systems_in_scope and len(functions_in_scope) < len(functions):
        pending.append(f"Escopo editorial da Matriz Funcional aplicado aos {len(systems_in_scope)} sistema(s) conciliado(s): {len(functions_in_scope)} de {len(functions)} registros foram incluídos. Sistemas sem base de usuários continuam documentados no diagnóstico.")
    observed, conflict_summary, user_conflicts = _observed_conflicts(matrix, contexts, rules, config["atividades_conflitantes"])
    sod_result = analyze_sod(matrix, contexts)
    if sod_result.template_conflicts:
        observed.extend(sod_result.template_conflicts)
        for impacted in sod_result.user_impacted:
            user_conflicts.append({"usuario": impacted["Usuário"], "sistema": impacted["Sistema"], "perfis": impacted["Perfis"], "id_conflito": impacted["Regra SoD"], "id_atividade_1": "", "atividade_1": "", "id_atividade_2": "", "atividade_2": ""})
    else:
        pending.append("Atividades Conflitantes não preenchidas: não foram localizadas atividades aderentes às regras conservadoras da metodologia fornecida.")
    if config.get("sat", {}).get("habilitado", False):
        sat = _sat_rows(matrix, config, pending)
    else:
        sat = sod_result.sat_candidates
        if sat:
            pending.append(f"SAT gerada com {len(sat)} atividade(s) pelos critérios metodológicos; todas permanecem A validar até a confirmação do workflow e das alçadas.")
        else:
            pending.append("SAT não preenchida: não foram identificadas ações explícitas aderentes aos critérios metodológicos; direitos indeterminados foram mantidos em lacunas.")
    # Divergência de perfil impede a entrega da matriz final. O diagnóstico, a
    # minuta e a análise preliminar continuam disponíveis para a tratativa.
    blocked_by_profile_divergence = bool(profile_divergence_issues)
    if blocked_by_profile_divergence:
        pending.append("Matriz Funcional final não gerada: há divergências de perfis entre as bases conciliadas. Regularize ou formalize os perfis antes de liberar o Excel oficial.")
    template = resolve_from_config(config_path, config["template_oficial"]["arquivo"])
    if not template.is_file():
        raise FileNotFoundError(f"Template oficial não encontrado: {template}")
    workbook_path = None
    if not blocked_by_profile_divergence:
        workbook_path = run_dir / _safe_output_name(config["template_oficial"].get("nome_saida", "Matriz_SoD_SAT_resultado.xlsx"), display_systems)
        export_template(template, workbook_path, config["template_oficial"], {"matriz_funcional": matrix, "atividades_conflitantes": observed, "sat": sat}, system_label=display_label)
    sod_analysis_path = run_dir / "Matriz_de_Conflitos_Consolidada.xlsx"
    _write_validation(
        validation_path,
        mappings,
        issues,
        blocked_by_profile_divergence=blocked_by_profile_divergence,
    )
    _write_csv(diagnostics_path, ["tipo", "sistema_usuarios", "sistema_funcionalidades", "perfil", "onde", "como_corrigir"], issues)
    _write_email_draft(email_draft_path, issues)
    export_analysis(sod_analysis_path, sod_result)
    _write_csv(system_mapping_path, ["sistema_usuarios", "sistema_funcionalidades", "metodo", "perfis_usuarios", "perfis_funcionalidades"], mappings)
    _write_csv(user_conflicts_path, ["usuario", "sistema", "perfis", "id_conflito", "id_atividade_1", "atividade_1", "id_atividade_2", "atividade_2"], user_conflicts)
    status = "BLOQUEADO_POR_DIVERGENCIA_DE_PERFIS" if blocked_by_profile_divergence else "GERADO_COM_PENDENCIAS" if pending else "GERADO_SEM_PENDENCIAS"
    log = {"run_id": run_id, "gerado_em_utc": datetime.now(timezone.utc).isoformat(), "status": status, "template_oficial": {"arquivo": str(template), "sha256": _sha256(template)}, "base_funcionalidades": functions_summary, "base_usuarios": users_summary, "atividades_conflitantes": rules_summary, "de_para_sistemas": mappings, "escopo_matriz_funcional": {"sistemas_conciliados": sorted(systems_in_scope), "sistema_exibido": display_label, "registros_entrada": len(functions), "registros_no_escopo": len(functions_in_scope)}, "referencia_cac": "Usada apenas como benchmark estrutural e de validação durante a evolução; não participa da decisão de conflitos da execução.", "transformacoes_aplicadas": ["Aliases de colunas configurados", "Apenas remoção de espaços no início/fim de valores mapeados", "Atividade preservada da fonte; texto editorial somente quando ausente", "Classificação SoD por ações, objetos e termos de negócio da própria fonte", "Fontes RAW preservadas sem alteração"], "inconsistencias": {"quantidade": len(issues), "sistema_perfil": len(reconciliation_issues), "qualidade_fonte": len(source_issues), "arquivo": str(diagnostics_path)}, "analise_sod_metodologica": {**sod_result.summary, "arquivo": str(sod_analysis_path), "metodo": "Regras conservadoras fornecidas pelo usuário; status padrão A validar"}, "reconciliacao": {**conflict_summary, "conflitos_encontrados": len(observed), "resultados_sat": len(sat), "registros_matriz": len(matrix), "regras_lidas": len(rules)}, "pendencias": pending}
    log_path.write_text(json.dumps(log, ensure_ascii=False, indent=2, default=str) + "\n", encoding="utf-8")
    pending_path.write_text(_pending_text(run_id, pending), encoding="utf-8")
    # A existência do Excel não elimina divergências de perfis. A minuta fica
    # disponível somente quando há reconciliação pendente a comunicar ao cliente.
    email_required = bool(profile_divergence_issues and contexts)
    return ProcessingResult(run_id, workbook_path, log_path, pending_path, validation_path, diagnostics_path, email_draft_path, sod_analysis_path, system_mapping_path, user_conflicts_path, email_required, status)

