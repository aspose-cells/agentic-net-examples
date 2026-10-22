// Title: Export a workbook to HTML while omitting hidden worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, hides one of its worksheets, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves the file as HTML, and then reads the HTML to confirm the hidden sheet data is absent. | Show a step‑by‑step C# example with Aspose.Cells that demonstrates hiding a worksheet, configuring HTML export options to exclude hidden sheets, saving the workbook, and programmatically verifying the omission.
// Common Searches: Aspose.Cells C# export workbook to HTML without hidden sheets | HtmlSaveOptions ExportHiddenWorksheet false example | how to hide a worksheet and exclude it from HTML output using Aspose.Cells | verify hidden worksheet content is not in generated HTML Aspose.Cells | C# Aspose.Cells HTML conversion options for visible worksheets only
// Tags: HtmlSaveOptions ExportHiddenWorksheet false | hide worksheet Aspose.Cells C# | export workbook to HTML Aspose.Cells | verify hidden sheet exclusion HTML | Aspose.Cells HTML conversion options

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a workbook with a visible and a hidden worksheet, configuring HtmlSaveOptions.ExportHiddenWorksheet = false, saving the workbook as HTML, and checking the generated HTML to ensure the hidden worksheet's content is not included.
class ExportHiddenWorksheetDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default first worksheet and add some data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "VisibleSheet";
            sheet1.Cells["A1"].PutValue("This is a visible sheet.");

            // Add a second worksheet and hide it
            int hiddenSheetIndex = workbook.Worksheets.Add();
            Worksheet hiddenSheet = workbook.Worksheets[hiddenSheetIndex];
            hiddenSheet.Name = "HiddenSheet";
            hiddenSheet.Cells["A1"].PutValue("This content should NOT appear in HTML.");
            hiddenSheet.IsVisible = false; // Hide the worksheet

            // Prepare HTML save options and exclude hidden worksheets from export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportHiddenWorksheet = false
                // Additional options can be set here if needed
            };

            // Define output folder and file
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputFolder);
            string htmlFilePath = Path.Combine(outputFolder, "Workbook.html");

            // Save the workbook as HTML using the options
            workbook.Save(htmlFilePath, htmlOptions);

            // Verify that the hidden worksheet content is not present in the generated HTML
            if (File.Exists(htmlFilePath))
            {
                string htmlContent = File.ReadAllText(htmlFilePath);
                bool hiddenContentFound = htmlContent.Contains("HiddenSheet") ||
                                          htmlContent.Contains("This content should NOT appear in HTML.");

                Console.WriteLine("HTML export completed.");
                Console.WriteLine(hiddenContentFound
                    ? "Verification FAILED: Hidden worksheet content was found in the HTML."
                    : "Verification PASSED: Hidden worksheet content is excluded from the HTML.");
            }
            else
            {
                Console.WriteLine("HTML file was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
