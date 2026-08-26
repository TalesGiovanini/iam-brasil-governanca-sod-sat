from tales.cli.main import build_parser


def test_cli_accepts_status():
    args = build_parser().parse_args(["status"])
    assert args.command == "status"
