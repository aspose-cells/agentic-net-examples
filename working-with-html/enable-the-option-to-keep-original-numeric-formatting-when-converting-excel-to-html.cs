using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.html";

                // Verify that the source Excel file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options (default options export data, formulas and formatting)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // Save the workbook as HTML using the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}