// Title: Create a macro‑free Excel report of cells using the Hyperlink theme color with Aspose.Cells for .NET (C#)
// Description: C# example that builds a new workbook, applies the Hyperlink theme color to sample cells, scans the used range, detects Font.ThemeColor = Hyperlink, and writes each cell's address and value to a separate report sheet. The workbook is saved macro‑free with auto‑fitted columns.
// Keywords: Aspose.Cells | C# | .NET | Hyperlink theme color | list cells by theme color | macro‑free Excel report | Font.ThemeColor filter | Excel cell audit | theme color detection | Aspose.Cells example
// Common Searches: Aspose.Cells list cells with Hyperlink theme color | C# generate report of hyperlink‑styled cells | How to find cells using Font.ThemeColor in Aspose.Cells | Macro‑free Excel file that lists theme colored cells | Aspose.Cells iterate cells and check ThemeColor
// Developer Intent: Generate a macro‑free Excel workbook that enumerates every cell whose font uses the Hyperlink theme color and records its address and value on a summary worksheet.
// Use Cases: Audit a worksheet to locate all hyperlink‑styled cells for compliance checks. | Create a concise summary sheet for reviewers who need to see only hyperlink‑colored data. | Export a clean, macro‑free Excel file for downstream processing or sharing. | Validate that theme‑based formatting has been applied consistently across a workbook. | Document cell addresses and contents for reporting or documentation purposes.
// AI Prompts: Write C# code using Aspose.Cells to scan a worksheet and list cells where Font.ThemeColor equals Hyperlink, outputting address and value to a new sheet. | Show how to modify the example to include separate columns for row index and column letter instead of the full cell address. | Explain how to generalize the solution to filter cells by any ThemeColorType (e.g., Accent1, Accent2) with a single parameter. | Provide steps to convert the macro‑free report into a CSV file using Aspose.Cells. | Describe how to add conditional formatting to the report sheet that highlights cells with empty values.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace HyperlinkThemeColorReport
{
    // C# example that builds a new workbook, applies the Hyperlink theme color to sample cells, scans the used range, detects Font.ThemeColor = Hyperlink, and writes each cell's address and value to a separate report sheet. The workbook is saved macro‑free with auto‑fitted columns.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet (source data)
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Data";

            // Add sample data and apply the Hyperlink theme color to some cells
            // Cell A1 - Hyperlink theme color
            Cell cellA1 = sourceSheet.Cells["A1"];
            cellA1.PutValue("Link 1");
            Style styleA1 = cellA1.GetStyle();
            styleA1.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0.0);
            cellA1.SetStyle(styleA1);

            // Cell B2 - Normal color
            Cell cellB2 = sourceSheet.Cells["B2"];
            cellB2.PutValue("Normal");
            // No special theme color applied

            // Cell C3 - Hyperlink theme color
            Cell cellC3 = sourceSheet.Cells["C3"];
            cellC3.PutValue("Link 2");
            Style styleC3 = cellC3.GetStyle();
            styleC3.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0.0);
            cellC3.SetStyle(styleC3);

            // Create a new worksheet for the report
            Worksheet reportSheet = workbook.Worksheets.Add("HyperlinkThemeReport");
            // Header
            reportSheet.Cells["A1"].PutValue("Cell Address");
            reportSheet.Cells["B1"].PutValue("Cell Value");

            int reportRow = 1; // zero‑based index; start after header

            // Iterate through all used cells in the source sheet
            int maxRow = sourceSheet.Cells.MaxDataRow;
            int maxCol = sourceSheet.Cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell curCell = sourceSheet.Cells[row, col];
                    // Skip empty cells
                    if (curCell.Type == CellValueType.IsNull) continue;

                    Style curStyle = curCell.GetStyle();

                    // Check if the font uses the Hyperlink theme color
                    if (curStyle.Font.ThemeColor != null &&
                        curStyle.Font.ThemeColor.ColorType == ThemeColorType.Hyperlink)
                    {
                        // Write the address and value to the report sheet
                        reportSheet.Cells[reportRow, 0].PutValue(curCell.Name);
                        reportSheet.Cells[reportRow, 1].PutValue(curCell.StringValue);
                        reportRow++;
                    }
                }
            }

            // Auto‑fit columns for better readability
            reportSheet.AutoFitColumns();

            // Save the workbook (macro‑free)
            workbook.Save("HyperlinkThemeColorReport.xlsx");
        }
    }
}
