using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class GetPaperWidth
    {
        static void Main()
        {
            // Path to the existing XLSX file
            string filePath = "input.xlsx";

            // Load the workbook from the file (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];

            // Retrieve the paper width (in inches) from the worksheet's PageSetup
            double paperWidthInInches = firstSheet.PageSetup.PaperWidth;

            // Output the result
            Console.WriteLine("Paper Width of the first worksheet: " + paperWidthInInches + " inches");
        }
    }
}