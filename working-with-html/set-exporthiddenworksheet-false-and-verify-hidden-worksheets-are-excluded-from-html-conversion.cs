// Title: How to exclude hidden worksheets when converting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML without hidden sheets | How to set ExportHiddenWorksheet to false in HtmlSaveOptions | Check that hidden worksheet content is omitted after HTML conversion | C# example for hiding a worksheet and skipping it in HTML output | Aspose.Cells HTML save options to exclude hidden worksheets
// Tags: HtmlSaveOptions ExportHiddenWorksheet false | Aspose.Cells hide worksheet during HTML export | C# workbook to HTML without hidden sheets | verify hidden worksheet exclusion in HTML output | Aspose.Cells HTML conversion settings

using System;
using System.IO;
using Aspose.Cells;

namespace ExportHiddenWorksheetDemo
{
    // The sample creates a workbook with one visible and one hidden worksheet, configures HtmlSaveOptions.ExportHiddenWorksheet = false, saves the workbook as HTML, and then reads the generated file to confirm that the hidden sheet's name and data are absent, reporting a pass/fail verification.
    class Program
    {
        static void Main(string[] args)
        {
            // Path for the generated HTML file
            string htmlOutputPath = "ExportHiddenWorksheetDemo.html";

            try
            {
                // ------------------------------------------------------------
                // 1. Create a workbook with two worksheets, one hidden.
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();

                // Visible worksheet
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("This is a visible sheet");
                visibleSheet.Cells["A2"].PutValue(DateTime.Now);

                // Hidden worksheet
                int hiddenIndex = workbook.Worksheets.Add();
                Worksheet hiddenSheet = workbook.Worksheets[hiddenIndex];
                hiddenSheet.Name = "HiddenSheet";
                hiddenSheet.Cells["A1"].PutValue("You should NOT see this in the HTML");
                hiddenSheet.Cells["A2"].PutValue(12345);
                hiddenSheet.IsVisible = false; // Hide the sheet

                // ------------------------------------------------------------
                // 2. Prepare HTML save options
                // ------------------------------------------------------------
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportHiddenWorksheet = false, // Do NOT export hidden worksheets
                    ExportGridLines = true,
                    ExportHeadings = true,
                    PageTitle = "ExportHiddenWorksheet Demo"
                };

                // ------------------------------------------------------------
                // 3. Save the workbook to HTML using the configured options
                // ------------------------------------------------------------
                workbook.Save(htmlOutputPath, htmlOptions);
                Console.WriteLine($"HTML file saved to: {Path.GetFullPath(htmlOutputPath)}");

                // ------------------------------------------------------------
                // 4. Verify that the hidden worksheet content is NOT present in the HTML
                // ------------------------------------------------------------
                if (File.Exists(htmlOutputPath))
                {
                    string htmlContent = File.ReadAllText(htmlOutputPath);

                    bool hiddenSheetFound = htmlContent.IndexOf("HiddenSheet", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                            htmlContent.IndexOf("You should NOT see this", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (hiddenSheetFound)
                    {
                        Console.WriteLine("Verification FAILED: Hidden worksheet content was found in the HTML output.");
                    }
                    else
                    {
                        Console.WriteLine("Verification PASSED: Hidden worksheet content is correctly excluded from the HTML output.");
                    }
                }
                else
                {
                    Console.WriteLine("Verification FAILED: HTML output file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during processing:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
