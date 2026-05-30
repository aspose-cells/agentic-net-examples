using System;
using System.IO;
using Aspose.Cells;

namespace HtmlToExcelWithHyperlinks
{
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for source HTML and resulting Excel files
            string htmlPath = "sample.html";
            string excelPath = "result.xlsx";

            // Verify that the HTML template exists before loading
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"Error: HTML file not found at '{htmlPath}'.");
                return;
            }

            try
            {
                // Load the HTML file into a Workbook.
                // Aspose.Cells automatically parses <a> tags and creates Hyperlink objects.
                Workbook workbook = new Workbook(htmlPath);

                // (Optional) Verify that hyperlinks were imported.
                // Iterate through the first worksheet's hyperlinks collection.
                Worksheet sheet = workbook.Worksheets[0];
                foreach (Hyperlink link in sheet.Hyperlinks)
                {
                    // Display the hyperlink address; row/column info is optional and omitted to avoid API version issues.
                    Console.WriteLine($"Address: {link.Address}");
                }

                // Save the workbook in Excel format. Hyperlinks remain clickable in the saved file.
                workbook.Save(excelPath, SaveFormat.Xlsx);

                Console.WriteLine($"HTML converted to Excel successfully. Output saved at: {excelPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}