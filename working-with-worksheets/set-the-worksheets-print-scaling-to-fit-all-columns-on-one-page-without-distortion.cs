// Title: Set Aspose.Cells worksheet print scaling to fit all columns on one page (C#)
// Description: Demonstrates how to configure a worksheet's PageSetup in Aspose.Cells for .NET so that the printed output fits all columns on a single page while allowing the row height to scale automatically. Includes sample data generation and saving the workbook.
// Keywords: Aspose.Cells print scaling | FitToPagesWide | FitToPagesTall | C# worksheet page setup | fit columns on one page | Aspose.Cells page layout | Excel print fit width | .NET Excel printing
// Common Searches: Aspose.Cells fit all columns on one printed page | C# set worksheet print scaling Aspose.Cells | FitToPagesWide = 1 Aspose.Cells example | How to print Excel sheet without distortion using Aspose.Cells | Aspose.Cells page setup fit width only
// Developer Intent: Configure the worksheet's print settings so the width fits on one page and the height adjusts automatically.
// Use Cases: Generating printable reports where horizontal data must stay on a single page. | Creating invoices or statements that need a consistent page width regardless of row count. | Exporting large data tables to Excel with automatic width scaling for clean printing.
// AI Prompts: Show C# code to set FitToPagesWide = 1 and FitToPagesTall = 0 in Aspose.Cells. | Explain how Aspose.Cells page setup scales columns without distorting rows. | Provide an Aspose.Cells example that fits all worksheet columns on one printed page.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintScalingDemo
{
    // Demonstrates how to configure a worksheet's PageSetup in Aspose.Cells for .NET so that the printed output fits all columns on a single page while allowing the row height to scale automatically. Includes sample data generation and saving the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some sample data to demonstrate the effect
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Set the page setup to fit all columns on one page.
            // FitToPagesWide = 1 means one page wide.
            // FitToPagesTall = 0 lets the height adjust automatically.
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.FitToPagesTall = 0;

            // Save the workbook
            workbook.Save("FitAllColumnsOnePage.xlsx");
        }
    }
}
