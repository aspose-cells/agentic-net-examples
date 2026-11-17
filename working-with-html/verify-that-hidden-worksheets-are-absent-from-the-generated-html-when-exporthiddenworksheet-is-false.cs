// Title: Verify hidden worksheets are omitted from HTML export when ExportHiddenWorksheet is false using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide worksheet from HTML export using ExportHiddenWorksheet | How to prevent hidden sheets from appearing in generated HTML with Aspose.Cells | C# example checking that hidden worksheet is not included in HTML output | ExportHiddenWorksheet property false usage in Aspose.Cells .NET | Validate HTML save options to exclude hidden worksheets in Aspose.Cells
// Tags: HtmlSaveOptions ExportHiddenWorksheet false C# | exclude hidden worksheets from HTML export Aspose.Cells | Aspose.Cells hidden sheet omission in HTML output | C# workbook HTML save without hidden sheets | verify hidden worksheet exclusion Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The sample creates a workbook with one visible and one hidden worksheet, sets HtmlSaveOptions.ExportHiddenWorksheet to false, saves the workbook as HTML to a memory stream, reads the generated HTML, and confirms that the hidden sheet name is absent while the visible sheet name is present, demonstrating how to exclude hidden worksheets from HTML export in Aspose.Cells for .NET.
class ExportHiddenWorksheetTest
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook wb = new Workbook();
        // Add a visible worksheet
        Worksheet visibleSheet = wb.Worksheets[0];
        visibleSheet.Name = "VisibleSheet";
        // Add a hidden worksheet
        int hiddenIndex = wb.Worksheets.Add();
        Worksheet hiddenSheet = wb.Worksheets[hiddenIndex];
        hiddenSheet.Name = "HiddenSheet";
        hiddenSheet.IsVisible = false; // Mark as hidden

        // Configure HTML save options: do NOT export hidden worksheets
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportHiddenWorksheet = false;

        // Save the workbook to a memory stream as HTML
        using (MemoryStream htmlStream = new MemoryStream())
        {
            wb.Save(htmlStream, htmlOptions);
            htmlStream.Position = 0;
            string htmlContent = new StreamReader(htmlStream, Encoding.UTF8).ReadToEnd();

            // Verify that the hidden worksheet name does NOT appear in the HTML
            bool hiddenSheetPresent = htmlContent.Contains("HiddenSheet");
            bool visibleSheetPresent = htmlContent.Contains("VisibleSheet");

            Console.WriteLine("Visible sheet present in HTML: " + visibleSheetPresent);
            Console.WriteLine("Hidden sheet present in HTML: " + hiddenSheetPresent);

            if (visibleSheetPresent && !hiddenSheetPresent)
            {
                Console.WriteLine("Test passed: hidden worksheets are absent from the generated HTML.");
            }
            else
            {
                Console.WriteLine("Test failed: hidden worksheets are incorrectly included.");
            }
        }
    }
}
