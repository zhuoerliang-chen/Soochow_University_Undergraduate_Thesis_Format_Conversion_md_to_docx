from __future__ import annotations

import argparse
import runpy
from pathlib import Path
from typing import Any

from word_format_config import DEFAULT_WORD_FORMAT, WordFormatConfig


def _load_word_format_config(config_path: Path) -> WordFormatConfig:
    data: dict[str, Any] = runpy.run_path(str(config_path))
    cfg = data.get("WORD_FORMAT") or data.get("CONFIG") or data.get("config")
    if isinstance(cfg, WordFormatConfig):
        return cfg
    raise ValueError(
        f"配置文件未提供 WordFormatConfig 实例（期望变量名 WORD_FORMAT / CONFIG / config）：{config_path}"
    )


def _cmd_convert(args: argparse.Namespace) -> int:
    from md_to_docx import convert_markdown_to_docx

    input_md = Path(args.input).resolve()
    output_docx = Path(args.output).resolve() if args.output else input_md.with_suffix(".docx")

    config = DEFAULT_WORD_FORMAT
    if args.config:
        config = _load_word_format_config(Path(args.config).resolve())

    saved = convert_markdown_to_docx(input_md, output_docx, config=config)
    print(str(saved))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="mdtodocx", description="Convert a Markdown document to DOCX with preset formatting.")
    sub = parser.add_subparsers(dest="command", required=True)

    p_convert = sub.add_parser("convert", help="Convert a Markdown file to DOCX.")
    p_convert.add_argument("-i", "--input", required=True, help="Input Markdown path.")
    p_convert.add_argument("-o", "--output", help="Output DOCX path. Default: same name as input.")
    p_convert.add_argument("--config", help="Path to a Python config file exporting WORD_FORMAT (WordFormatConfig).")
    p_convert.set_defaults(func=_cmd_convert)

    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    return int(args.func(args))


if __name__ == "__main__":
    raise SystemExit(main())
