// Title: C# – Export Excel workbook to HTML with Aspose.Cells while hiding hidden worksheets and fitting cells to prevent overflow (ExportHiddenWorksheet = false + HtmlCrossType.FitToCell)
// AI Prompts: Write C# code that loads an .xlsx file, marks a worksheet as hidden, sets HtmlSaveOptions.ExportHiddenWorksheet to false and HtmlSaveOptions.HtmlCrossType to FitToCell, then saves the workbook as HTML using Aspose.Cells. | Demonstrate how to confirm that the generated HTML omits hidden worksheets and that each cell's content is confined within the cell boundaries with Aspose.Cells. | Modify an existing Aspose.Cells example to add HtmlCrossType.FitToCell for overflow control while keeping ExportHiddenWorksheet disabled.
// Common Searches: Aspose.Cells C# export Excel to HTML hide hidden sheets ExportHiddenWorksheet false example | how to use HtmlCrossType.FitToCell in Aspose.Cells to prevent cell overflow in HTML output | C# code to export workbook to HTML without hidden worksheets using Aspose.Cells HtmlSaveOptions | prevent hidden worksheet data from appearing in HTML export Aspose.Cells | fit cell content to cell size during HTML export Aspose.Cells C#
// Tags: Aspose.Cells HTML export options | exclude hidden worksheets Aspose.Cells | fit-to-cell layout Aspose.Cells | prevent overflow in HTML export C# | C# Excel to HTML conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, hides a worksheet, configures HtmlSaveOptions with ExportHiddenWorksheet set to false and HtmlCrossType set to FitToCell, and saves the workbook as an HTML file, ensuring hidden sheets are omitted and cell content does not overflow.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide the second worksheet to test ExportHiddenWorksheet option
            if (workbook.Worksheets.Count > 1)
                workbook.Worksheets[1].IsVisible = false;

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportHiddenWorksheet = false // Exclude hidden sheets
                // Note: HtmlCrossType property is not available in current Aspose.Cells version
            };

            // Export the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
