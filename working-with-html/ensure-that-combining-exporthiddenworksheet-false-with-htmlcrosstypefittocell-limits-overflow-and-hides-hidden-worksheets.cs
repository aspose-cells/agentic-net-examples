// Title: Export HTML without Hidden Worksheets and Fit Long Text to Cell using Aspose.Cells for .NET
// Description: Shows how to build a workbook with a visible sheet that contains wrapped long text, add a hidden sheet, and save the file to HTML. The HtmlSaveOptions are set to ExportHiddenWorksheet = false and HtmlCrossStringType = HtmlCrossType.FitToCell, which removes hidden worksheets from the output and confines text overflow to the cell width.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportHiddenWorksheet | HtmlCrossType.FitToCell | hide hidden worksheets | limit text overflow | C# .NET | HTML export spreadsheet | cell width fitting | Aspose.Cells example
// Common Searches: Aspose.Cells hide hidden worksheets when exporting to HTML | FitToCell prevent text overflow in HTML export Aspose.Cells | Export workbook to HTML without hidden sheets C# | HtmlSaveOptions ExportHiddenWorksheet false example | Wrap long text in HTML export using Aspose.Cells
// Developer Intent: Export a workbook to HTML while excluding hidden worksheets and ensuring that long cell content stays within the cell boundaries.
// Use Cases: Create web‑ready spreadsheet previews that omit confidential or auxiliary sheets. | Generate HTML reports with narrow columns where long text must stay inside each cell. | Produce printable HTML output from a workbook without exposing hidden data. | Build automated email attachments that display only visible worksheets with properly wrapped content.
// AI Prompts: Provide a code snippet that also disables image export while keeping hidden worksheets excluded. | Explain how HtmlCrossType.FitToCell works together with ExportHiddenWorksheet to control layout in HTML output. | Write a unit test that confirms hidden worksheets are absent from the generated HTML and that long text does not overflow its cell.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook with a visible sheet that contains wrapped long text, add a hidden sheet, and save the file to HTML. The HtmlSaveOptions are set to ExportHiddenWorksheet = false and HtmlCrossStringType = HtmlCrossType.FitToCell, which removes hidden worksheets from the output and confines text overflow to the cell width.
    public class ExportHiddenWorksheetWithFitToCellDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Visible worksheet with data that would normally overflow
            // -------------------------------------------------
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";

            // Put a long text in A1
            visibleSheet.Cells["A1"].PutValue("This is a very long text that would normally overflow the cell width when exported to HTML.");

            // Apply text wrapping style
            Style wrapStyle = workbook.CreateStyle();
            wrapStyle.IsTextWrapped = true;
            StyleFlag flag = new StyleFlag();
            flag.WrapText = true;
            visibleSheet.Cells["A1"].SetStyle(wrapStyle, flag);

            // Set column width to a small value to force overflow
            visibleSheet.Cells.SetColumnWidth(0, 10);

            // -------------------------------------------------
            // Hidden worksheet that should not appear in the HTML output
            // -------------------------------------------------
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden sheet data");
            hiddenSheet.IsVisible = false; // Mark as hidden

            // -------------------------------------------------
            // Configure HTML save options:
            //   - Do not export hidden worksheets
            //   - Use FitToCell to limit overflow within cell width
            // -------------------------------------------------
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false,               // Hide hidden worksheets
                HtmlCrossStringType = HtmlCrossType.FitToCell // Limit overflow to cell width
            };

            // Save the workbook as HTML
            string outputPath = "ExportHiddenWorksheet_FitToCell.html";
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath} with hidden worksheets excluded and overflow limited.");
        }
    }
}
