using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CenterOnPageDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add some sample data to visualize the effect
            worksheet.Cells["A1"].PutValue("This worksheet will be centered horizontally and vertically when printed.");
            worksheet.Cells["A2"].PutValue("CenterOnPage = true");

            // Access the PageSetup object of the worksheet
            PageSetup pageSetup = worksheet.PageSetup;

            // Center the sheet horizontally and vertically on the printed page
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = true;

            // (Optional) Define a print area to see the centering effect clearly
            pageSetup.PrintArea = "A1:B2";

            // Save the workbook to a file
            workbook.Save("CenterOnPageDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CenterOnPageDemo.Run();
        }
    }
}