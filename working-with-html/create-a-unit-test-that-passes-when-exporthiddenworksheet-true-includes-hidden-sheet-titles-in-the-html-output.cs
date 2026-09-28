// Title: Create a C# unit test that verifies ExportHiddenWorksheet=true includes hidden worksheet titles in Aspose.Cells HTML output
// AI Prompts: Generate a C# test method that adds a hidden worksheet, saves the workbook with HtmlSaveOptions.ExportHiddenWorksheet enabled, and asserts that the hidden sheet name appears inside an <h2> tag in the resulting HTML. | Write a C# example that uses a MemoryStream to capture Aspose.Cells HTML output and checks for the presence of the hidden worksheet title when ExportHiddenWorksheet is set to true.
// Common Searches: how to verify hidden worksheets are exported in Aspose.Cells HTML using C# | Aspose.Cells HtmlSaveOptions ExportHiddenWorksheet unit test example | C# check hidden sheet title in generated HTML from Aspose.Cells | validate hidden worksheet visibility in HTML export Aspose.Cells | unit testing Aspose.Cells HTML export of hidden sheets
// Tags: Aspose.Cells ExportHiddenWorksheet HTML option | C# test hidden worksheet HTML export | memory stream Aspose.Cells HTML capture | detect hidden worksheet name in HTML | hidden worksheet visibility in HTML export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a workbook with one visible and one hidden worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet to true, saves the workbook to a MemoryStream as HTML, reads the HTML string, and throws an exception if the hidden sheet title (<h2>HiddenSheet</h2>) is not found, effectively serving as a unit test for the ExportHiddenWorksheet feature.
    public class ExportHiddenWorksheetTests
    {
        // Entry point for the example
        public static void Main()
        {
            try
            {
                ExecuteTest();
                Console.WriteLine("Test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Replicates the original NUnit test logic without NUnit dependencies
        public static void ExecuteTest()
        {
            // Create a new workbook
            var workbook = new Workbook();

            // Add a visible worksheet
            var visibleSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            visibleSheet.Name = "VisibleSheet";

            // Add a hidden worksheet
            var hiddenSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            hiddenSheet.Name = "HiddenSheet";
            hiddenSheet.IsVisible = false; // Hide the sheet

            // Prepare HTML save options
            var htmlOptions = new HtmlSaveOptions
            {
                // Export hidden worksheets
                ExportHiddenWorksheet = true
            };

            // Save the workbook to a memory stream as HTML
            using (var htmlStream = new MemoryStream())
            {
                try
                {
                    workbook.Save(htmlStream, htmlOptions);
                }
                catch (Exception saveEx)
                {
                    throw new InvalidOperationException("Failed to save workbook as HTML.", saveEx);
                }

                htmlStream.Position = 0;

                // Read the generated HTML as a string
                string htmlContent;
                using (var reader = new StreamReader(htmlStream))
                {
                    htmlContent = reader.ReadToEnd();
                }

                // Verify that the hidden sheet title appears in the HTML output
                // The sheet title is rendered inside a <h2> tag by default
                if (!htmlContent.Contains("<h2>HiddenSheet</h2>"))
                {
                    throw new InvalidOperationException(
                        "HTML output does not contain the title of the hidden worksheet when ExportHiddenWorksheet is true.");
                }
            }
        }
    }
}
