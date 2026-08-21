// Title: C# – Convert Excel workbook to HTML with BestFit layout and export hidden worksheets using Aspose.Cells
// Description: Creates a workbook with a visible and a hidden sheet, configures HtmlSaveOptions (PresentationPreference = true, ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false) and saves the entire workbook as a single HTML file on the desktop.
// Keywords: Aspose.Cells | C# | .NET | HtmlSaveOptions | PresentationPreference | BestFit | ExportHiddenWorksheet | ExportActiveWorksheetOnly | convert Excel to HTML | hidden worksheets | save workbook as HTML
// Common Searches: Aspose.Cells export hidden sheets to HTML C# | HtmlSaveOptions PresentationPreference BestFit example | Save entire Excel workbook as HTML with hidden worksheets | C# convert workbook to HTML Aspose.Cells | How to include hidden worksheets in HTML export using Aspose
// Developer Intent: Generate a single HTML file that represents the whole Excel workbook, applying the BestFit presentation style and including any hidden worksheets.
// Use Cases: Render a complete workbook (including hidden configuration sheets) in a web portal for reporting. | Create an HTML snapshot of all worksheets for archival or email distribution while preserving layout. | Provide an HTML preview in a SaaS application where hidden sheets contain metadata required for calculations.
// AI Prompts: Show how to attach a custom CSS stylesheet to the HTML output while still exporting hidden worksheets. | Explain how to export only selected worksheets with PresentationPreference set to BestFit. | Give a step‑by‑step guide to split a large workbook into separate HTML files per worksheet, ensuring hidden sheets are included.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook with a visible and a hidden sheet, configures HtmlSaveOptions (PresentationPreference = true, ExportHiddenWorksheet = true, ExportActiveWorksheetOnly = false) and saves the entire workbook as a single HTML file on the desktop.
    public class ConvertWorkbookToHtml
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Add sample data to the first (visible) worksheet
            Worksheet visibleSheet = workbook.Worksheets[0];
            visibleSheet.Name = "VisibleSheet";
            visibleSheet.Cells["A1"].PutValue("Visible Data");

            // Add a hidden worksheet with some data
            Worksheet hiddenSheet = workbook.Worksheets.Add("HiddenSheet");
            hiddenSheet.Cells["A1"].PutValue("Hidden Data");
            hiddenSheet.IsVisible = false; // Mark the sheet as hidden

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Enable presentation preference for a more beautiful layout (BestFit)
                PresentationPreference = true,

                // Ensure hidden worksheets are exported
                ExportHiddenWorksheet = true,

                // Export all worksheets (including hidden) rather than only the active one
                ExportActiveWorksheetOnly = false
            };

            // Define output HTML file path
            string outputPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "WorkbookExport.html");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved to HTML at: {outputPath}");
        }
    }
}
