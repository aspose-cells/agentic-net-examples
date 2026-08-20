// Title: Aspose.Cells C# – Macro‑Free Report of Cells Using the Hyperlink Theme Color
// Description: Creates a workbook, applies the Hyperlink theme color to selected cells, scans the used range, records each cell address and its hyperlink URL (if any) on a separate report sheet, auto‑fits columns, and saves a macro‑free XLSX file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel | hyperlink theme color | list themed cells | macro‑free report | detect hyperlink style | theme color detection | Excel automation | cell style scanning
// Common Searches: list cells with hyperlink theme color Aspose.Cells .NET | generate Excel report of themed hyperlink cells without macros | find cells styled with Hyperlink theme color using C# | Aspose.Cells detect font ThemeColor Hyperlink | create summary of hyperlink‑styled cells in Excel
// Developer Intent: Produce a macro‑free Excel file that enumerates every cell whose font uses the Hyperlink theme color and includes any associated hyperlink address.
// Use Cases: Audit spreadsheets to verify that only intended cells use the Hyperlink theme color. | Export a concise index of themed hyperlink cells for documentation or UI generation. | Validate workbook styling before publishing to ensure compliance with branding guidelines.
// AI Prompts: Write C# code with Aspose.Cells that scans a worksheet and outputs a report of cells using the Hyperlink theme color together with their hyperlink URLs. | Provide a method that returns a DataTable of cell addresses and hyperlink addresses where the font ThemeColor equals Hyperlink. | Explain how to extend the example to also capture cells colored with a custom RGB value that matches the default hyperlink color.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace HyperlinkThemeColorReport
{
    // Creates a workbook, applies the Hyperlink theme color to selected cells, scans the used range, records each cell address and its hyperlink URL (if any) on a separate report sheet, auto‑fits columns, and saves a macro‑free XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Add sample cells with hyperlinks and apply the Hyperlink theme color to the font
            // Cell A1 - hyperlink with Hyperlink theme color
            sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");
            Cell a1 = sheet.Cells["A1"];
            a1.PutValue("Example Link");
            Style styleA1 = a1.GetStyle();
            // Set font to use the Hyperlink theme color
            styleA1.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0.0);
            a1.SetStyle(styleA1);

            // Cell B2 - regular text (no theme color)
            Cell b2 = sheet.Cells["B2"];
            b2.PutValue("Regular Text");
            // No special theme color applied

            // Cell C3 - hyperlink but default style (not theme color)
            sheet.Hyperlinks.Add("C3", 1, 1, "https://www.aspose.com");
            Cell c3 = sheet.Cells["C3"];
            c3.PutValue("Aspose Link");
            // Keep default style

            // Cell D4 - apply Hyperlink theme color without a hyperlink
            Cell d4 = sheet.Cells["D4"];
            d4.PutValue("Styled Text");
            Style styleD4 = d4.GetStyle();
            styleD4.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0.0);
            d4.SetStyle(styleD4);

            // Create a report worksheet
            Worksheet report = workbook.Worksheets.Add("Report");
            // Header
            report.Cells[0, 0].PutValue("Cell Address");
            report.Cells[0, 1].PutValue("Hyperlink Address (if any)");

            int reportRow = 1;

            // Iterate through all used cells in the data worksheet
            int maxRow = sheet.Cells.MaxDisplayRange.RowCount;
            int maxCol = sheet.Cells.MaxDisplayRange.ColumnCount;

            for (int row = 0; row < maxRow; row++)
            {
                for (int col = 0; col < maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Skip empty cells
                    if (cell.Type == CellValueType.IsNull) continue;

                    Style cellStyle = cell.GetStyle();
                    // Check if the font uses the Hyperlink theme color
                    if (cellStyle.Font.ThemeColor != null &&
                        cellStyle.Font.ThemeColor.ColorType == ThemeColorType.Hyperlink)
                    {
                        // Record the cell address
                        string address = cell.Name; // e.g., "A1"
                        report.Cells[reportRow, 0].PutValue(address);

                        // If the cell also contains a hyperlink, retrieve its address
                        string hyperlinkAddress = string.Empty;
                        foreach (Hyperlink link in sheet.Hyperlinks)
                        {
                            if (link.Area.StartRow == row && link.Area.StartColumn == col)
                            {
                                hyperlinkAddress = link.Address;
                                break;
                            }
                        }
                        report.Cells[reportRow, 1].PutValue(hyperlinkAddress);
                        reportRow++;
                    }
                }
            }

            // Auto-fit columns for better readability
            report.AutoFitColumns();

            // Save the workbook (macro‑free)
            workbook.Save("HyperlinkThemeColorReport.xlsx");
        }
    }
}
