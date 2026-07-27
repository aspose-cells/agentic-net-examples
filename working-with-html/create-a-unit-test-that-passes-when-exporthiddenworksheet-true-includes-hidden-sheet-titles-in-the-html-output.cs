// Title: C# Unit Test – Verify ExportHiddenWorksheet Shows Hidden Sheet Titles in HTML (Aspose.Cells)
// Description: Creates a workbook with a visible and a hidden worksheet, saves it to HTML using HtmlSaveOptions.ExportHiddenWorksheet = true, reads the generated file and asserts that the hidden sheet name appears, then removes the temporary file.
// Keywords: Aspose.Cells | ExportHiddenWorksheet | HTML export | C# unit test | hidden worksheet | HtmlSaveOptions | automated testing | MSTest | NUnit | xUnit
// Common Searches: Aspose.Cells unit test hidden sheet HTML export | ExportHiddenWorksheet true not showing hidden sheet | how to assert hidden worksheet name in HTML output | C# test for Aspose.Cells HTML save options | verify hidden worksheets are exported with Aspose.Cells
// Developer Intent: Write an automated test that confirms hidden worksheets are included and their titles appear when a workbook is saved to HTML with ExportHiddenWorksheet enabled.
// Use Cases: Regression test to ensure hidden sheets are not omitted after library updates. | Validate compliance reports that require hidden worksheet metadata in HTML exports. | Continuous‑integration check for correct HtmlSaveOptions configuration across projects.
// AI Prompts: Generate an MSTest method that creates a workbook with a hidden sheet, saves it to HTML with ExportHiddenWorksheet = true, and asserts the hidden sheet name exists in the output. | Write an NUnit test case for Aspose.Cells that verifies hidden worksheets are exported to HTML and cleans up the temporary file. | Provide an xUnit example that checks the HTML content contains the hidden worksheet title after saving with HtmlSaveOptions.ExportHiddenWorksheet set to true.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTests
{
    // Creates a workbook with a visible and a hidden worksheet, saves it to HTML using HtmlSaveOptions.ExportHiddenWorksheet = true, reads the generated file and asserts that the hidden sheet name appears, then removes the temporary file.
    public class HtmlExportHiddenWorksheetDemo
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and add data to the default (visible) sheet
                Workbook workbook = new Workbook();
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("Visible Data");

                // Add a hidden worksheet with some data
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Hidden Data");
                hiddenSheet.IsVisible = false; // Mark the sheet as hidden

                // Configure HTML save options to export hidden worksheets
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportHiddenWorksheet = true,          // Ensure hidden sheets are exported
                    ExportActiveWorksheetOnly = false      // Export the whole workbook
                };

                // Save the workbook to a temporary HTML file
                string htmlFilePath = Path.Combine(Path.GetTempPath(), "ExportHiddenWorksheetTest.html");
                workbook.Save(htmlFilePath, saveOptions);

                // Verify that the hidden sheet title appears in the HTML output
                if (File.Exists(htmlFilePath))
                {
                    string htmlContent = File.ReadAllText(htmlFilePath);
                    bool containsHidden = htmlContent.Contains("HiddenSheet");
                    Console.WriteLine(containsHidden
                        ? "Success: Hidden worksheet title is present in the HTML output."
                        : "Failure: Hidden worksheet title is NOT present in the HTML output.");
                }
                else
                {
                    Console.WriteLine("Error: HTML file was not created.");
                }

                // Clean up the temporary file
                if (File.Exists(htmlFilePath))
                {
                    File.Delete(htmlFilePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception occurred: {ex.Message}");
            }
        }
    }
}
