// Title: HTML to Excel in C# with CSS border mapping – Aspose.Cells
// Description: Loads an HTML file into an Aspose.Cells Workbook, optionally translates CSS border names to CellBorderType values via a dictionary, applies the mapped line styles to every used cell, and saves the result as an XLSX workbook.
// Keywords: Aspose.Cells | HTML to Excel conversion | C# | .NET | CSS border mapping | CellBorderType | Excel cell borders | LoadOptions Html | Workbook.Save Xlsx
// Common Searches: convert html file to xlsx using Aspose.Cells | map css border styles to Excel borders C# | Aspose.Cells custom border mapping example | load html workbook and preserve borders | apply CellBorderType from css in Aspose.Cells
// Developer Intent: Load an HTML document, replace CSS border definitions with matching Excel border styles, and generate a styled XLSX file.
// Use Cases: Transform HTML reports into Excel workbooks while keeping original border thicknesses. | Process HTML tables that store border style names in comments and reflect them in Excel cells. | Create a reusable utility that reads CSS border values, maps them to Aspose.Cells enums, and applies consistent styling across worksheets.
// AI Prompts: Write C# code that parses CSS border attributes from HTML and maps them to Aspose.Cells CellBorderType enums for all cells. | Show how to extend the CSS‑to‑CellBorder dictionary with custom names and apply different line styles to individual cell sides. | Explain a fallback strategy when a CSS border style is missing from the mapping dictionary during HTML‑to‑Excel conversion.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Alias to avoid ambiguity with System.Range
using AsposeRange = Aspose.Cells.Range;

// Loads an HTML file into an Aspose.Cells Workbook, optionally translates CSS border names to CellBorderType values via a dictionary, applies the mapped line styles to every used cell, and saves the result as an XLSX workbook.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Paths for source HTML and destination Excel files
        string htmlFile = "sample.html";
        string excelFile = "converted.xlsx";

        try
        {
            // Ensure the HTML source file exists
            if (!File.Exists(htmlFile))
            {
                Console.WriteLine($"Error: HTML file '{htmlFile}' not found.");
                return;
            }

            // Load the HTML file into a workbook using LoadOptions (HTML format)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
            Workbook workbook = new Workbook(htmlFile, loadOptions);

            // ------------------------------------------------------------
            // OPTIONAL: Custom mapping of CSS border styles to Excel borders
            // ------------------------------------------------------------
            var cssToCellBorder = new Dictionary<string, CellBorderType>(StringComparer.OrdinalIgnoreCase)
            {
                { "thin", CellBorderType.Thin },
                { "medium", CellBorderType.Medium },
                { "dashed", CellBorderType.Dashed },
                { "dotted", CellBorderType.Dotted },
                { "double", CellBorderType.Double },
                { "hair", CellBorderType.Hair },
                { "mediumDashed", CellBorderType.MediumDashed },
                { "mediumDashDot", CellBorderType.MediumDashDot },
                { "mediumDashDotDot", CellBorderType.MediumDashDotDot },
                { "slantedDashDot", CellBorderType.SlantedDashDot } // corrected enum name
            };

            // Iterate over all used cells and adjust border line styles if needed
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the range that actually contains data
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                int startRow = usedRange.FirstRow;
                int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                int startColumn = usedRange.FirstColumn;
                int endColumn = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        Style style = cell.GetStyle();

                        // Placeholder: suppose we stored the original CSS border style name
                        // in the cell's comment. In real scenarios you would extract the CSS
                        // value from the HTML parser.
                        string? cssBorder = cell.Comment?.Note;

                        if (!string.IsNullOrEmpty(cssBorder) && cssToCellBorder.TryGetValue(cssBorder, out CellBorderType borderType))
                        {
                            // Apply the mapped border type to all four sides
                            style.Borders[BorderType.LeftBorder].LineStyle = borderType;
                            style.Borders[BorderType.RightBorder].LineStyle = borderType;
                            style.Borders[BorderType.TopBorder].LineStyle = borderType;
                            style.Borders[BorderType.BottomBorder].LineStyle = borderType;
                            cell.SetStyle(style);
                        }
                    }
                }
            }

            // Save the workbook as an Excel file (XLSX)
            workbook.Save(excelFile, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlFile}' has been converted to Excel file '{excelFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
