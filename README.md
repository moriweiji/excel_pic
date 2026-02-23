# excel-pic

Episode Excel generator (core CLI stage).

## Environment

- Package manager: `uv`
- Python: `>=3.11`

## Install dependencies

```bash
uv sync
```

## Run

```bash
uv run excel-pic -i "第21集" -w "第21集/剧情分镜（21-30）.docx" --exe-dir .
```

## Run GUI

```bash
uv run excel-pic-gui
```

Options:

- `-i, --images-dir` 图片源目录
- `-w, --word` Word 路径（可在任意目录）
- `--strict` 严格模式（坏图等问题直接中断）
- `--prefix-text` 临时覆盖前缀模板文本
- `--exe-dir` 模拟 exe 同级目录（默认当前目录）

## Output layout

```text
<exe_dir>/data/
  config/prompt_prefix.txt
  export/<输入目录名>/
    *.png
    <集数>集.xlsx
  logs/
    *_<输入目录名>.log
    *_<输入目录名>_report.json
```

## Notes

- Excel images are embedded (not external links), so the export folder can be zipped and shared.
- Source assets are not modified.
- This public repository intentionally excludes local docs and sample assets.

## Release (Windows exe)

Tag and push a version to trigger GitHub Release build:

```bash
git tag v0.1.0
git push origin v0.1.0
```

The workflow will attach `excel-pic-gui.exe` and `SHA256SUMS.txt` to the GitHub Release.

