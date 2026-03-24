# Excel Template Cell Style Creator

A .NET 8.0 console application that interactively builds an Excel file containing a visual gallery of custom cell styles. Each style is rendered as a "Sample Text" row so you can see exactly how it will look in your own workbooks.

---

## Prerequisites

| Requirement | Details |
|-------------|---------|
| .NET 8.0 SDK | [Download](https://dotnet.microsoft.com/download/dotnet/8.0) |
| Output directory | `C:\temp\` must exist before running |

---

## How to run

```bash
cd ExcelTemplateCellStyleCreator
dotnet run
```

The application writes `ExcelStyleTemplate.xlsx` to `C:\temp\`. Any existing file at that path is deleted before a new one is created.

---

## Interactive prompts

For each style the application asks for the following. Press **Enter** to accept the shown default.

| Prompt | Accepted values | Default |
|--------|----------------|---------|
| Font name | Any installed font family name | `Calibri` |
| Font size | Positive number | `11` |
| Font color | 6-digit hex (e.g. `FF0000`) | `000000` (black) |
| Bold | `y` / `n` | `n` |
| Italic | `y` / `n` | `n` |
| Background color | 6-digit hex | `FFFFFF` (white) |
| Border sides | Any combination of `l` `r` `t` `b` | `lrtb` (all sides) |
| Configure alignment? | `y` / `n` | `n` |

If you choose `y` for alignment, three additional prompts appear:

| Prompt | Accepted values |
|--------|----------------|
| Horizontal alignment | `l` (left), `c` (center), `r` (right) |
| Vertical alignment | `t` (top), `c` (center), `b` (bottom) |
| Wrap text | `y` / `n` |

After each style, you are asked whether to add another. Duplicate styles are silently skipped.

---

## Output

The generated workbook contains a single sheet named **Styles**:

- **Column A** — `StyleIndex Id = N` label for each style
- **Column B** — `Sample Text` rendered with the defined style

Rows are spaced two apart for readability. Grid lines are hidden.

---

## Reference colors

| Color | Hex |
|-------|-----|
| Red | `FF0000` |
| Green | `00FF00` |
| Blue | `0000FF` |
| Yellow | `FFFF00` |
| Black | `000000` |
| White | `FFFFFF` |

---

## Project structure

| File | Purpose |
|------|---------|
| `Program.cs` | Entry point; main loop, document creation, row insertion |
| `StyleManager.cs` | Manages OpenXML stylesheet collections; get-or-create helpers for fonts, fills, borders, cell formats |
| `UserInputValidator.cs` | Validates and normalizes all console input; loops until valid values are provided |
| `StyleDefaults.cs` | Carries last-used style values so subsequent prompts pre-fill sensibly |
| `LocalizationHelper.cs` | Returns German or English strings based on the system culture |
| `FileManager.cs` | Deletes the output file before re-creation |
