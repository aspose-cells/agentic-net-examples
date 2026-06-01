using System;
using Aspose.Cells;

namespace AsposeCellsFitToPageDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // (Optional) Add some sample data to demonstrate the effect
            worksheet.Cells["A1"].PutValue("Sample Data 1");
            worksheet.Cells["A2"].PutValue("Sample Data 2");
            worksheet.Cells["B1"].PutValue("Sample Data 3");
            worksheet.Cells["B2"].PutValue("Sample Data 4");

            // Set the page setup to fit the entire sheet on a single page
            worksheet.PageSetup.FitToPagesWide = 1; // Fit to 1 page wide
            worksheet.PageSetup.FitToPagesTall = 1; // Fit to 1 page tall

            // Save the workbook to a file
            workbook.Save("FitToSinglePage.xlsx");

            Console.WriteLine("Workbook saved with FitToPagesWide=1 and FitToPagesTall=1.");
        }
    }
}