// Title: Write a C# unit test using Aspose.Cells to verify that ExportHiddenWorksheet=true adds hidden worksheet titles to the generated HTML
// AI Prompts: Generate a C# test method (NUnit/MSTest) that creates a workbook with a hidden sheet, sets HtmlSaveOptions.ExportHiddenWorksheet to true, saves the workbook to a MemoryStream as HTML, reads the HTML string, and asserts that the hidden sheet name is present. | Produce a self‑contained C# unit test that uses Aspose.Cells to export a workbook to HTML with hidden worksheets included, then checks the output for the hidden worksheet title and fails the test if it is missing.
// Common Searches: Aspose.Cells unit test to check ExportHiddenWorksheet flag in HTML export | C# verify hidden worksheet appears in HTML when using HtmlSaveOptions.ExportHiddenWorksheet | How to assert hidden sheet title in HTML output with Aspose.Cells | Testing Aspose.Cells HTML export of hidden worksheets in .NET | Example of using MemoryStream to validate hidden worksheet inclusion in HTML
// Tags: Aspose.Cells HTML export hidden worksheet verification | C# HtmlSaveOptions ExportHiddenWorksheet unit test | assert hidden sheet title in generated HTML | memory stream HTML output validation Aspose.Cells | unit testing Aspose.Cells HTML export | hidden worksheet inclusion HTML Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // The example creates a workbook with one visible and one hidden worksheet, enables HtmlSaveOptions.ExportHiddenWorksheet, saves the workbook to a MemoryStream as HTML, reads the HTML content, and asserts that the hidden worksheet title is present, providing a pass/fail result.
    public class ExportHiddenWorksheetTests
    {
        public void Run()
        {
            try
            {
                // Create a new workbook with two worksheets
                var workbook = new Workbook();

                // Rename the default first sheet to a visible sheet name
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";

                // Add a second worksheet and hide it
                int hiddenSheetIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
                hiddenSheet.Name = "HiddenSheet";
                hiddenSheet.IsVisible = false; // Hide the worksheet

                // Prepare HTML save options and enable exporting of hidden worksheets
                var htmlOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true // Include hidden sheets in the HTML output
                };

                // Export the workbook to HTML using a memory stream
                using (var htmlStream = new MemoryStream())
                {
                    workbook.Save(htmlStream, htmlOptions);
                    htmlStream.Position = 0;

                    // Read the generated HTML as a string
                    string htmlContent;
                    using (var reader = new StreamReader(htmlStream))
                    {
                        htmlContent = reader.ReadToEnd();
                    }

                    // Verify that the hidden sheet title appears in the HTML output
                    if (htmlContent.Contains("HiddenSheet"))
                    {
                        Console.WriteLine("Test passed: Hidden sheet title found in HTML.");
                    }
                    else
                    {
                        Console.WriteLine("Test failed: Hidden sheet title not found in HTML.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main()
        {
            var test = new ExportHiddenWorksheetTests();
            test.Run();
        }
    }
}
