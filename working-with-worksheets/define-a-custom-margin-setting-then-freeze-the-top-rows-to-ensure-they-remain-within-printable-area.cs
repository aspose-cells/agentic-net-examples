using System;
using Aspose.Cells;

namespace AsposeCellsMarginAndFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // ----- Custom margin settings (centimeters) -----
            PageSetup pageSetup = worksheet.PageSetup;
            pageSetup.TopMargin = 2.0;      // 2 cm top margin
            pageSetup.BottomMargin = 1.5;   // 1.5 cm bottom margin
            pageSetup.LeftMargin = 1.0;     // 1 cm left margin
            pageSetup.RightMargin = 1.0;    // 1 cm right margin

            // Ensure the first two rows repeat on each printed page
            pageSetup.PrintTitleRows = "$1:$2";

            // ----- Freeze the top two rows -----
            // Freeze panes at cell A3, freezing 2 rows and 0 columns
            worksheet.FreezePanes("A3", 2, 0);

            // Save the workbook
            workbook.Save("MarginAndFreezeDemo.xlsx");
        }
    }
}