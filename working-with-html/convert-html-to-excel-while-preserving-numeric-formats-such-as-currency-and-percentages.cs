// Title: How to convert an HTML file to Excel in C# with Aspose.Cells while preserving currency and percentage formatting
// AI Prompts: Load an HTML document using HtmlLoadOptions with ConvertNumericData = true, iterate through all cells, detect strings that end with '%' or start with '$', convert them to numeric values, apply the built‑in percentage (Number = 10) or currency (Number = 164) style, and save the workbook as XLSX. | Extend the cell‑processing loop to also handle numeric strings that use parentheses for negatives or custom symbols, and assign the appropriate built‑in number format via Aspose.Cells styling.
// Common Searches: Aspose.Cells C# convert HTML to XLSX keep currency symbols | preserve percentage formatting when loading HTML with HtmlLoadOptions | ConvertNumericData option example for HTML to Excel conversion | apply built‑in number formats to cells after importing HTML in Aspose.Cells | C# change string "$1,200" to numeric currency cell in Excel
// Tags: HtmlLoadOptions ConvertNumericData | HTML to XLSX numeric formatting | currency cell style Aspose.Cells | percentage number format Aspose.Cells | C# cell value conversion after HTML import

using System;
using Aspose.Cells;

// The example loads an HTML file with HtmlLoadOptions (ConvertNumericData enabled), scans each cell to find string representations of percentages and US dollar amounts, converts those strings to numeric values, applies the built‑in percentage (10) or currency (164) number formats, and saves the result as an XLSX workbook using Aspose.Cells for .NET.
class HtmlToExcelConverter
{
    static void Main()
    {
        // Load the HTML file with options that preserve numeric data
        var loadOptions = new HtmlLoadOptions(LoadFormat.Html)
        {
            // Convert numeric strings (e.g., "$123.45", "67%") to numeric cells
            ConvertNumericData = true
        };
        var workbook = new Workbook("input.html", loadOptions);

        // After loading, fine‑tune cells that represent currency or percentages
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            for (int col = 0; col <= cells.MaxDataColumn; col++)
            {
                Cell cell = cells[row, col];

                // Process only string cells that may contain formatted numbers
                if (cell.Type == CellValueType.IsString)
                {
                    string text = cell.StringValue.Trim();

                    // Handle percentages (e.g., "45%")
                    if (text.EndsWith("%"))
                    {
                        if (double.TryParse(text.TrimEnd('%'), out double percentValue))
                        {
                            cell.PutValue(percentValue / 100); // Store as decimal
                            Style style = cell.GetStyle();
                            style.Number = 10; // Built‑in percentage format
                            cell.SetStyle(style);
                        }
                    }
                    // Handle currency values prefixed with "$" (e.g., "$1,234.56")
                    else if (text.StartsWith("$"))
                    {
                        // Remove currency symbol and any grouping commas
                        string numericPart = text.Substring(1).Replace(",", "");
                        if (double.TryParse(numericPart, out double currencyValue))
                        {
                            cell.PutValue(currencyValue);
                            Style style = cell.GetStyle();
                            style.Number = 164; // Built‑in currency format with "$"
                            cell.SetStyle(style);
                        }
                    }
                }
            }
        }

        // Save the workbook as an Excel file
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
