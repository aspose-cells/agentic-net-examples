// Title: C# – Export Hidden Worksheets to HTML with Aspose.Cells and Verify Content
// Description: Creates a workbook with a visible and a hidden sheet, configures HtmlSaveOptions to export hidden worksheets, saves the file as HTML, and programmatically confirms that the hidden sheet name and data are present in the output.
// Keywords: Aspose.Cells C# export hidden worksheet HTML | HtmlSaveOptions ExportHiddenWorksheet true | include hidden sheets in HTML export | verify hidden sheet content Aspose.Cells | ExportActiveWorksheetOnly false
// Common Searches: Aspose.Cells export hidden worksheet to HTML C# example | HtmlSaveOptions ExportHiddenWorksheet property usage | how to include hidden sheets when saving as HTML with Aspose.Cells | validate hidden sheet data in generated HTML
// Developer Intent: Enable ExportHiddenWorksheet, save the workbook as HTML, and ensure hidden worksheets are rendered in the resulting file.
// Use Cases: Produce an HTML preview that shows data from hidden tabs for audit reports. | Create web‑based documentation where supplemental information resides on hidden sheets. | Automate CI checks that confirm hidden worksheet content is correctly exported.
// AI Prompts: Generate C# code using Aspose.Cells to export a workbook to HTML while including hidden worksheets and validate the output. | Explain the impact of ExportHiddenWorksheet and ExportActiveWorksheetOnly on HTML conversion in Aspose.Cells. | Suggest ways to extend the verification step to check hidden sheet formatting and styles in the HTML.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportHiddenWorksheetDemo
{
    // Creates a workbook with a visible and a hidden sheet, configures HtmlSaveOptions to export hidden worksheets, saves the file as HTML, and programmatically confirms that the hidden sheet name and data are present in the output.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ----- Visible worksheet -----
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // ----- Hidden worksheet -----
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options to export hidden worksheets
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = true,          // Ensure hidden sheets are exported
                ExportActiveWorksheetOnly = false      // Export the whole workbook
            };

            // Define output HTML file path
            string outputHtmlPath = Path.Combine(Environment.CurrentDirectory, "WorkbookWithHiddenSheet.html");

            // Save the workbook as HTML
            workbook.Save(outputHtmlPath, saveOptions);
            Console.WriteLine($"Workbook saved to HTML at: {outputHtmlPath}");

            // ----- Verification -----
            // Read the generated HTML file
            string htmlContent = File.ReadAllText(outputHtmlPath);

            // Check if the hidden sheet's data appears in the HTML
            bool hiddenDataFound = htmlContent.Contains("Hidden Data");
            bool hiddenSheetNameFound = htmlContent.Contains("HiddenSheet");

            // Output verification result
            if (hiddenDataFound && hiddenSheetNameFound)
            {
                Console.WriteLine("Verification succeeded: Hidden worksheet content is included in the HTML.");
            }
            else
            {
                Console.WriteLine("Verification failed: Hidden worksheet content is NOT found in the HTML.");
            }
        }
    }
}
