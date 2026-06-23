using System;
using Aspose.Cells;

namespace AsposeCellsPrintScalingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle rule)
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

            // Set print scaling: fit all columns on one page (wide) and let height adjust automatically
            // FitToPagesWide = 1 means one page wide.
            // FitToPagesTall = 0 disables height constraint, allowing rows to span multiple pages.
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.FitToPagesTall = 0;

            // Ensure scaling is based on FitToPages properties, not percent scaling
            worksheet.PageSetup.IsPercentScale = false;

            // Save the workbook (lifecycle rule)
            workbook.Save("FitAllColumnsOnePage.xlsx");
        }
    }
}