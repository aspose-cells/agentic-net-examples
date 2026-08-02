// Title: Set narrow print margins in an Excel worksheet with Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, and configures the PageSetup margins to 0.5 cm (≈0.2 in) on all sides using both centimeter‑based and inch‑based properties, then saves the file as NarrowMarginsDemo.xlsx to maximize printable area.
// Keywords: Aspose.Cells | C# print margins | Excel PageSetup | narrow margins | centimeter margins | inch margins | maximize printable area | set worksheet margins programmatically
// Common Searches: Aspose.Cells set worksheet margins | C# narrow print margins Excel | reduce Excel print margins with Aspose | PageSetup LeftMarginInch vs LeftMargin | set margins in centimeters Aspose.Cells
// Developer Intent: Apply minimal left, right, top, and bottom margins to a worksheet to increase the printable area when generating Excel files.
// Use Cases: Print a multi‑page report that fits more rows per page by using very small margins. | Generate invoices or receipts where the content should occupy almost the entire page. | Design a flyer or booklet layout in Excel with maximum usable space and minimal white borders.
// AI Prompts: Provide C# code using Aspose.Cells to set all page margins to 0.3 cm and export the workbook as a PDF. | Explain how LeftMargin and LeftMarginInch properties are linked in Aspose.Cells and which one overrides the other. | Show how to read the current worksheet margin values in both centimeters and inches from PageSetup.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintMarginsDemo
{
    // Creates a workbook, adds sample data, and configures the PageSetup margins to 0.5 cm (≈0.2 in) on all sides using both centimeter‑based and inch‑based properties, then saves the file as NarrowMarginsDemo.xlsx to maximize printable area.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data so the printed page has content
            sheet.Cells["A1"].PutValue("Demo of narrow print margins");
            for (int row = 2; row <= 20; row++)
            {
                sheet.Cells[$"A{row}"].PutValue($"Row {row - 1}");
                sheet.Cells[$"B{row}"].PutValue(row * 10);
            }

            // Set narrow margins (values are in centimeters)
            // 0.5 cm ≈ 0.2 inches – very small margins to maximize printable area
            sheet.PageSetup.LeftMargin = 0.5;
            sheet.PageSetup.RightMargin = 0.5;
            sheet.PageSetup.TopMargin = 0.5;
            sheet.PageSetup.BottomMargin = 0.5;

            // Optionally set the same values using the inch‑based properties
            // (demonstrates both APIs; they refer to the same underlying settings)
            sheet.PageSetup.LeftMarginInch = 0.2;
            sheet.PageSetup.RightMarginInch = 0.2;
            sheet.PageSetup.TopMarginInch = 0.2;
            sheet.PageSetup.BottomMarginInch = 0.2;

            // Save the workbook (lifecycle rule: save)
            workbook.Save("NarrowMarginsDemo.xlsx");
        }
    }
}
