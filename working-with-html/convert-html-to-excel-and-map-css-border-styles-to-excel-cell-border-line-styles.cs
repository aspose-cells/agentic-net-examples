// Title: C# – Convert HTML to Excel and Translate CSS Border Styles to Excel Cell Borders with Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, walks every populated cell, converts CSS‑derived thin and medium borders to medium and thick Excel line styles, and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells HTML to Excel | C# CSS border mapping | Excel cell border style conversion | Load HTML workbook .NET | Save workbook as XLSX | CellBorderType Thin to Medium | CellBorderType Medium to Thick | HTML table to spreadsheet
// Common Searches: how to convert html to excel with aspose.cells c# | map css border thickness to excel cell borders | change cell border line style after loading html | asp.net convert html tables to xlsx preserving borders | c# iterate cells to adjust border style aspose
// Developer Intent: Load an HTML document into a Workbook, adjust each cell’s border line type to reflect the original CSS thickness, and export the modified workbook as an XLSX file.
// Use Cases: Transform web‑generated tables into Excel while upgrading thin CSS borders for clearer presentation. | Standardize border thickness across imported HTML reports to maintain consistent spreadsheet formatting. | Automate batch conversion of multiple HTML files, applying custom CSS‑to‑Excel border mappings before saving.
// AI Prompts: Write C# code that loads an HTML file with Aspose.Cells, iterates all cells, replaces CellBorderType.Thin with Medium and Medium with Thick, then saves as .xlsx. | Show how to extend the border‑mapping logic to support solid, dashed, and double CSS borders when converting HTML to Excel using Aspose.Cells. | Explain how to access the original CSS border properties of an HTML cell via LoadOptions in Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an HTML file into an Aspose.Cells Workbook, walks every populated cell, converts CSS‑derived thin and medium borders to medium and thick Excel line styles, and saves the result as an XLSX workbook.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Input HTML file path (must exist)
        string htmlPath = "sample.html";

        // Output Excel file path
        string excelPath = "converted.xlsx";

        // -------------------------------------------------
        // Load the HTML file into a Workbook (load lifecycle)
        // -------------------------------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
        Workbook workbook = new Workbook(htmlPath, loadOptions);

        // -------------------------------------------------
        // Map CSS border styles to Excel cell border line styles
        // (simple example mapping based on existing line style)
        // -------------------------------------------------
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    Style style = cell.GetStyle();
                    bool styleChanged = false;

                    // Iterate over all four border sides
                    foreach (BorderType borderSide in new[] {
                        BorderType.LeftBorder,
                        BorderType.RightBorder,
                        BorderType.TopBorder,
                        BorderType.BottomBorder })
                    {
                        CellBorderType currentLine = style.Borders[borderSide].LineStyle;

                        // Example mapping:
                        // - Thin  -> Medium (assume original CSS was "solid")
                        // - Medium -> Thick  (assume original CSS was "double")
                        // Adjust as needed for your specific CSS-to-Excel mapping.
                        if (currentLine == CellBorderType.Thin)
                        {
                            style.Borders[borderSide].LineStyle = CellBorderType.Medium;
                            styleChanged = true;
                        }
                        else if (currentLine == CellBorderType.Medium)
                        {
                            style.Borders[borderSide].LineStyle = CellBorderType.Thick;
                            styleChanged = true;
                        }
                    }

                    if (styleChanged)
                    {
                        cell.SetStyle(style);
                    }
                }
            }
        }

        // -------------------------------------------------
        // Save the workbook as an Excel file (save lifecycle)
        // -------------------------------------------------
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}
