// Title: Localize numeric cell formatting to French (fr-FR) culture in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, sets the workbook CultureInfo to a given locale, scans all used cells, and applies a built‑in number format (e.g., format ID 4) to cells that contain numeric values while keeping their existing styles intact. | Create a reusable method ApplyLocalizedNumberFormat(string inputPath, string outputPath, string cultureCode) that loads a workbook, configures CultureInfo, iterates the used range, detects numeric cells via IsNumericValue, and updates only the number‑format portion of each cell's style. | Generate a snippet that uses a regular expression to confirm that a cell's DisplayStringValue contains digits before assigning the localized number format, ensuring text cells remain unchanged.
// Common Searches: Aspose.Cells C# set French number format for all numeric cells in an existing workbook | How to preserve cell formatting while changing number format with Aspose.Cells .NET | Detect numeric values in Excel cells and apply culture‑specific format using Aspose.Cells | Apply built‑in number format ID 4 to numeric cells after setting workbook CultureInfo | Iterate used range and update only number format of cells in Aspose.Cells C#
// Tags: localized number format assignment Aspose.Cells | numeric cell identification Aspose.Cells C# | style flag number‑format update Aspose.Cells | workbook culture configuration Aspose.Cells | used‑range traversal Aspose.Cells

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an .xlsx workbook, sets its CultureInfo to a target locale (e.g., fr-FR), iterates the used range, identifies numeric cells, checks the displayed string for numeric characters, and applies built‑in number format ID 4 using a StyleFlag so only the number‑format part changes, then saves the workbook.
class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the desired culture for number formatting (e.g., French - France)
        workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

        // Work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the used range of the worksheet
        int maxRow = sheet.Cells.MaxDataRow;
        int maxCol = sheet.Cells.MaxDataColumn;

        // Regular expression to identify numeric content in the formatted string
        Regex numericRegex = new Regex(@"[\d.,]+", RegexOptions.Compiled);

        // Iterate through all cells in the used range
        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                Cell cell = sheet.Cells[row, col];

                // Skip empty cells
                if (cell == null || cell.Type == CellValueType.IsNull)
                    continue;

                // Process only cells that contain a numeric value (int, double, DateTime, etc.)
                if (cell.IsNumericValue)
                {
                    // Get the cell's current display string (already formatted by Excel style)
                    string displayValue = cell.DisplayStringValue;

                    // If the display string contains a numeric pattern, apply a localized format
                    if (numericRegex.IsMatch(displayValue))
                    {
                        // Retrieve the existing style
                        Style style = cell.GetStyle();

                        // Use a built‑in number format that will be localized automatically.
                        // Format 4 corresponds to "#,##0.00" (two decimal places).
                        style.Number = 4;

                        // Apply only the number format part of the style to preserve other formatting
                        StyleFlag flag = new StyleFlag();
                        flag.NumberFormat = true;

                        // Set the updated style back to the cell
                        cell.SetStyle(style, flag);
                    }
                }
            }
        }

        // Save the workbook with the localized number formats applied
        workbook.Save("output.xlsx");
    }
}
