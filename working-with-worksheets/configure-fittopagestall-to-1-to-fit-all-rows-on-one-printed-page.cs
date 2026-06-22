using System;
using Aspose.Cells;

namespace AsposeCellsFitToPagesTallDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the sheet with enough rows to require scaling
            for (int i = 0; i < 200; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    worksheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Get the PageSetup object for the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Define the print area (optional, but clarifies what will be printed)
            pageSetup.PrintArea = "A1:E200";

            // Set FitToPagesTall to 1 so all rows fit on a single printed page
            pageSetup.FitToPagesTall = 1;

            // According to the remarks, set FitToPagesWide to 0 to let the width adjust automatically
            pageSetup.FitToPagesWide = 0;

            // Save the workbook to demonstrate the configuration
            workbook.Save("FitToPagesTallDemo.xlsx");

            Console.WriteLine("Workbook saved with FitToPagesTall set to 1.");
        }
    }
}