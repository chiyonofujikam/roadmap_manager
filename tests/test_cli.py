from roadmap.helpers import get_parser


def test_cli_create_normal():
    parser = get_parser()
    args = parser.parse_args(["create", "--way", "normal"])
    assert args.action == "create"
    assert args.way == "normal"


def test_cli_delete_force():
    parser = get_parser()
    args = parser.parse_args(["delete", "--force"])
    assert args.force is True
