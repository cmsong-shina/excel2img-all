# excel2img-all

> [!NOTE] This codes written by AI.

> [!IMPORTANT] This only works on Windows with desktop Microsoft Excel installed.

Export every sheet from Excel workbooks to PNG.

This use Microsoft Excel to render workbook formatting.

## Requirements

- Windows with desktop Microsoft Excel installed.
- Python 3.12 or later.
- `uv` to run the project and install its `Pillow` and `pywin32` dependencies.

Supported input formats are `.xlsx`, `.xlsm`, `.xlsb`, and `.xls`. The `--dry-run`
option lists matching files and planned outputs without starting Excel and also
works on other platforms.

## Quick start

Run these commands from the project directory. Quote glob patterns so the exporter
can expand them, including recursive `**` patterns.

```powershell
# Export all sheets from matching workbooks
uvx --from "git+https://github.com/cmsong-shina/excel2img-all.git@master" `
    excel2img-all "some.xlsx" "other.xlsx"

# Search recursively and group output folders under images
uvx --from "git+https://github.com/cmsong-shina/excel2img-all.git@master" `
    excel2img-all "reports/**/*.xlsx" --output-dir images
```

## Options

| Option                        | Description                                                                                                  |
| ----------------------------- | ------------------------------------------------------------------------------------------------------------ |
| `-o`, `--output-dir DIR`      | Group workbook output folders under `DIR`. By default, each folder is created beside its source workbook.    |
| `--gridlines`                 | Show cell gridlines. Hidden by default; explicitly styled cell borders remain visible.                       |
| `--zoom N`                    | Excel view zoom from 10 to 400. Default: `100`.                                                              |
| `--retries N`                 | Maximum capture attempts per worksheet. Default: `3`, including the first attempt.                           |
| `--clipboard-timeout SECONDS` | Time to wait for Excel readiness, clipboard access, or a clipboard image at each step. Default: `8` seconds. |
| `--dry-run`                   | Preview matching files, available content ranges, and output folders without starting Excel.                 |

## Attribution and license

The original [excel2img](https://github.com/glexey/excel2img), written by Alexey
Gaydyukov, exports Excel ranges as images. This project builds on its Excel
`CopyPicture` approach with standalone script execution, glob matching, batch
exports, and clipboard retries.

Licensed under the [Apache License 2.0](LICENSE).
