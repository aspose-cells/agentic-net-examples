using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HtmlSaveOptionsPageTitleDemo
    {
        public static void Run()
        {
            // Path to the source Excel file
            string excelPath = "sample.xlsx";

            // Verify that the source file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            try
            {
                // Load the workbook from the file
                Workbook workbook = new Workbook(excelPath);

                // Create HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                // Set the page title to the workbook's file name (including extension)
                saveOptions.PageTitle = Path.GetFileName(excelPath);

                // Define the output HTML file path
                string htmlPath = "sample.html";

                // Save the workbook as HTML with the custom page title
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"HTML file saved to '{htmlPath}' with page title '{saveOptions.PageTitle}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}