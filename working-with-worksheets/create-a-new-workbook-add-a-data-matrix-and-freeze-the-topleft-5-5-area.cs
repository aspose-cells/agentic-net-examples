// Title: C# – Create a Workbook, Fill a 10×10 Matrix, and Freeze the Top‑Left 5×5 Area with Aspose.Cells
// Description: This example shows how to instantiate a Workbook, populate cells A1:J10 with sample values, freeze the first five rows and five columns using Worksheet.FreezePanes, and save the result as FreezeTopLeft5x5.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# freeze panes | FreezePanes top left area | create workbook Aspose.Cells | populate 10x10 matrix Aspose.Cells | .NET Excel freeze rows columns | Aspose.Cells example FreezeTopLeft5x5
// Common Searches: Aspose.Cells freeze first 5 rows and columns | C# code to freeze top left area in Excel | How to use Worksheet.FreezePanes in Aspose.Cells | Create and freeze a 5x5 area with Aspose.Cells .NET | Sample code for freezing panes in Aspose.Cells
// Developer Intent: Generate a new Excel workbook, insert a 10×10 data matrix, freeze the top‑left 5×5 region, and save the file.
// Use Cases: Financial dashboards where header rows and columns must stay visible while scrolling. | Data‑entry templates that keep row/column identifiers fixed for easier navigation. | Large matrix exports where the top‑left quadrant serves as a reference guide.
// AI Prompts: Write C# code with Aspose.Cells to create a workbook, add a 10×10 matrix, freeze the top‑left 5×5 area, and save it. | Explain each parameter of Worksheet.FreezePanes for freezing rows and columns in Aspose.Cells. | Show how to adjust the example to freeze only the first five rows or only the first five columns.

using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // This example shows how to instantiate a Workbook, populate cells A1:J10 with sample values, freeze the first five rows and five columns using Worksheet.FreezePanes, and save the result as FreezeTopLeft5x5.xlsx using Aspose.Cells for .NET.
    class FreezeTopLeftArea
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate a 10x10 data matrix with sample values
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze the top‑left 5 × 5 area
            // Parameters: row index, column index, number of frozen rows, number of frozen columns
            sheet.FreezePanes(5, 5, 5, 5);

            // Save the workbook (lifecycle rule: Save)
            workbook.Save("FreezeTopLeft5x5.xlsx");
        }
    }
}
