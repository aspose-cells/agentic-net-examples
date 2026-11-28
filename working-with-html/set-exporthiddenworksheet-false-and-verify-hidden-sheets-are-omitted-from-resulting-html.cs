// Title: Export a workbook to HTML without hidden worksheets using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML excluding hidden sheets | HtmlSaveOptions ExportHiddenWorksheet false example in .NET | how to verify hidden worksheet is not included in generated HTML using Aspose.Cells | save Excel file as HTML without hidden worksheets using Aspose.Cells for .NET | C# code to omit hidden worksheets when converting Excel to HTML
// Tags: HtmlSaveOptions ExportHiddenWorksheet false | export hidden worksheets exclusion Aspose.Cells | HTML conversion of Excel without hidden sheets | C# verify hidden sheet omission in HTML output | Aspose.Cells workbook to HTML without hidden worksheets

using System;
using System.IO;
using Aspose.Cells;

namespace ExportHiddenWorksheetDemo
{
    // The sample creates a workbook with one visible and one hidden worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves the workbook as an HTML file, and then reads the generated HTML to confirm that the hidden worksheet name is absent while the visible worksheet name is present.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for output HTML
            string outputHtmlPath = "ExportHiddenWorksheetDemo.html";

            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();

                // First (visible) worksheet
                Worksheet visibleSheet = workbook.Worksheets[0];
                visibleSheet.Name = "VisibleSheet";
                visibleSheet.Cells["A1"].PutValue("ID");
                visibleSheet.Cells["B1"].PutValue("Name");
                visibleSheet.Cells["A2"].PutValue(1);
                visibleSheet.Cells["B2"].PutValue("Alice");
                visibleSheet.Cells["A3"].PutValue(2);
                visibleSheet.Cells["B3"].PutValue("Bob");

                // Second worksheet that will be hidden
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.Cells["A1"].PutValue("Secret");
                hiddenSheet.Cells["A2"].PutValue("TopSecret");
                // Hide the worksheet
                hiddenSheet.IsVisible = false;

                // 2. Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do NOT export hidden worksheets
                    ExportHiddenWorksheet = false
                };

                // 3. Save the workbook to HTML
                workbook.Save(outputHtmlPath, htmlOptions);
                Console.WriteLine($"Workbook saved to HTML file: {outputHtmlPath}");

                // 4. Verify that the hidden sheet content is omitted
                if (File.Exists(outputHtmlPath))
                {
                    string htmlContent = File.ReadAllText(outputHtmlPath);

                    bool hiddenSheetFound = htmlContent.Contains(hiddenSheet.Name);
                    if (hiddenSheetFound)
                    {
                        Console.WriteLine("Verification FAILED: Hidden worksheet name was found in the HTML output.");
                    }
                    else
                    {
                        Console.WriteLine("Verification PASSED: Hidden worksheet was correctly omitted from the HTML output.");
                    }

                    // Additional check: ensure visible sheet data appears
                    bool visibleSheetFound = htmlContent.Contains(visibleSheet.Name);
                    if (visibleSheetFound)
                    {
                        Console.WriteLine("Visible worksheet is present in the HTML output as expected.");
                    }
                    else
                    {
                        Console.WriteLine("Warning: Visible worksheet name was not found in the HTML output.");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: Output HTML file '{outputHtmlPath}' was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
