// Title: Export a workbook to HTML in C# with hidden worksheets excluded and cell overflow fitted to cell using Aspose.Cells
// Common Searches: Aspose.Cells hide hidden worksheets when saving to HTML C# | How to prevent cell overflow in HTML export with Aspose.Cells .NET | Example of HtmlSaveOptions ExportHiddenWorksheet false and HtmlCrossType FitToCell | C# Aspose.Cells HTML export limit text overflow and exclude hidden sheets | Using HtmlCrossType.FitToCell with ExportHiddenWorksheet in Aspose.Cells
// Tags: Aspose.Cells exclude hidden sheets HTML export | Aspose.Cells FitToCell HTML overflow handling | C# HtmlSaveOptions ExportHiddenWorksheet configuration | HTML conversion cell overflow control Aspose.Cells | Aspose.Cells workbook to HTML C# example

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // The example creates a workbook with a visible worksheet containing long text and a hidden worksheet, configures HtmlSaveOptions to set ExportHiddenWorksheet to false and HtmlCrossType to FitToCell, then saves the workbook as an HTML file, ensuring hidden sheets are omitted and cell overflow is constrained to the cell width.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first (default) worksheet and add some data that would normally overflow
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "VisibleSheet";
                sheet1.Cells["A1"].PutValue("This is a very long text that will overflow the cell width if not handled properly.");

                // Create a style using the recommended factory method
                Style style = workbook.CreateStyle();
                style.HorizontalAlignment = TextAlignmentType.Center;
                sheet1.Cells["A1"].SetStyle(style);

                // Add a second worksheet and hide it
                Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
                hiddenSheet.IsVisible = false; // Mark worksheet as hidden
                hiddenSheet.Cells["A1"].PutValue("Hidden data that should not appear in the HTML output.");

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Do not export hidden worksheets
                    ExportHiddenWorksheet = false
                    // Additional options can be set here if needed
                };

                // Save the workbook as HTML using the configured options
                string outputPath = "ExportedWorkbook.html";
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully exported to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
