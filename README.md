# Factory Image SKU Handoff

**0.1.1-beta.1 · Free MIT beta · Windows and Python 3.12 tested**

Prepare embedded factory workbook images for a website builder without changing original text SKUs or image bytes. You declare the worksheet layout; the tool records candidate image-to-row mappings and flags ambiguous cases. A person must confirm that each picture actually depicts its candidate SKU.

## Install locally

For Codex, place the entire extracted package at `.agents/skills/factory-image-sku-handoff` inside your project. For direct CLI use, any folder you control is suitable. From that package folder, use your Python 3.12 executable, shown below as `python`:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip --isolated install --index-url https://pypi.org/simple --only-binary=:all: --require-hashes -r requirements-win-py312.lock
```

This creates a local environment and installs four separately licensed dependencies. It does not change global agent settings. The hash-locked binary wheels are for 64-bit Windows and CPython 3.12. `requirements.txt` also records exact versions; other platforms or Python versions have not been accepted. See [third-party notices](THIRD_PARTY_NOTICES.md).

For Codex, invoke `$factory-image-sku-handoff` or select it through `/skills`. Local changes are normally detected automatically; restart Codex if the skill does not appear. See the [official local skill instructions](https://learn.chatgpt.com/docs/build-skills). The complete package, including scripts, dependencies and examples, belongs under the same skill folder. This setup uses your project's local skill directory and does not change global configuration. Other agents can follow the supplied `SKILL.md` explicitly, and the CLI can be used directly.

## Run the original synthetic example

From the package folder, run this one command with a fresh output name:

```powershell
.\.venv\Scripts\python.exe -B scripts/handoff.py examples/products.xlsx --sheet Products --sku-column A --image-columns B --first-data-row 2 --output demo-handoff
```

`demo-handoff` must not already exist. The example has two explicitly sized product rows, one PNG and one JPEG. Its SKUs are exactly `000072` and ` 000073-蓝 `, including the second value's spaces. Both drawings meet the declared row rules. No real factory data or product photographs are included.

The expected result is exit 0, two assets, two occurrences and two rows. Compare `demo-handoff/manifest.json` and its asset bytes with `examples/expected/`; the manifest and asset hashes should match exactly. The manifest records the input workbook SHA-256. Expected outputs are evidence for this synthetic example only.

## Use a workbook

Supply all four layout choices explicitly: sheet name, SKU column, image start column(s) and first data row. For two possible image columns use, for example, `--image-columns B,D`. Multiple images on the same row still require review. Use a new output directory under an existing parent. The input workbook remains unchanged.

| Exit | Output | Meaning |
|---|---|---|
| 0 | `eligible` manifest and original supported image assets | All observations and SKU rows meet the declared layout rules, with no unresolved issues. Semantic ownership still needs human confirmation. |
| 2 | `review` manifest and original supported image assets | At least one ambiguity, unsupported layout or missing image needs review. Keep its reason codes with the handoff. |
| 1 | JSON error on stderr, no output directory | A package, input, resource, parser, argument or output boundary rejected the operation. |

The JSON manifest contains exact text SKU values, types, number formats, cells, drawing anchor evidence, every inventoried image occurrence, asset SHA-256 hashes, and explicit review reasons. Asset filenames are fixed hashes; SKU text is never used as a path. Identical image bytes share an asset but keep separate occurrences. Unsupported formats are inventoried without an exported asset.

## Supported scope and limits

- Drawing-based PNG/JPEG images are supported. One-cell anchors need an explicit positive row height and must fit vertically in that row; their start column must be declared. Horizontal extent is recorded, without claiming horizontal containment.
- Two-cell anchors require observable containment in one image cell. Nonzero horizontal offsets require review; nonzero vertical offsets need an explicit row height. Absolute, spanning and unsupported anchors require review.
- Blank, numeric, Boolean, date, error and formula SKUs require review. Text whitespace, leading zeroes and Unicode are preserved exactly. Duplicates, missing or multiple images, hidden rows/columns/sheets and merged candidate cells require review.
- WPS `DISPIMG`, Excel richData/in-cell images, VML, groups and other unsupported objects are reviewed or rejected. The package inventory is reconciled against parsed drawings to avoid silent omissions.
- [The contract](CONTRACT.md) defines byte, ZIP expansion, XML element, cell, merged-range, image pixel, row, column and occurrence limits. Unsafe ZIP names, duplicate members or cell coordinates, malformed row coordinates, active content, external relationships, DTD/entity declarations and parser mismatches are rejected before output.
- Nonstandard workbook/worksheet XML filename extensions and opaque relationships other than image or printer-settings parts are unsupported. The parser uses version-sensitive openpyxl models, so the exact tested dependency versions matter.

This is local handoff evidence. It does not create hosted image URLs, upload products, generate photographs, write a storefront import file, or prove real factory accuracy or customer demand.

## Privacy and data boundaries

The CLI uses no network, model, account or telemetry. Dependency installation accesses PyPI; workbook processing is local. It never evaluates formulas or executes macros, hyperlinks or workbook instructions.

Manifests intentionally contain source SKU text and image evidence. Treat those files as sensitive when the workbook is sensitive, and treat their strings as untrusted data when rendering them elsewhere. The tool does not anonymize SKU values. Authorize any external recipient or upload separately. Original workbooks are not copied into the output.

## License

Original code, documentation and synthetic examples in this beta are available under the [MIT license](LICENSE). Third-party distributions keep their own licenses. The beta is supplied without warranty. Synthetic tests establish only the tested local behavior; review actual business material before use.
