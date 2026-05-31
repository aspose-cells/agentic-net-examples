using System;
using Aspose.Cells;

namespace AsposeCellsShowZeroValuesExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Disable displaying zero values for the first worksheet (or all worksheets as needed)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.DisplayZeros = false;
            }

            // Configure HTML save options (default options are sufficient for this task)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as an HTML page
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("Workbook has been saved to HTML with zero values hidden.");
        }
    }
}