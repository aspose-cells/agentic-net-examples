// Title: Export Workbook to HTML with Original Sheet Order – Aspose.Cells C# Example (no custom IFilePathProvider)
// Description: Demonstrates how to save an Aspose.Cells workbook as HTML while keeping the original worksheet sequence. The sample creates three sheets, uses HtmlSaveOptions with ExportActiveWorksheetOnly = false and ExportHiddenWorksheet = true, and relies on the default IFilePathProvider so the generated HTML files follow the workbook's sheet order.
// Keywords: Aspose.Cells | HTML export | worksheet order | C# | HtmlSaveOptions | ExportActiveWorksheetOnly | ExportHiddenWorksheet | default IFilePathProvider | multi‑sheet HTML | save workbook as HTML
// Common Searches: Aspose.Cells export workbook to HTML keep sheet order | HTML export all worksheets Aspose.Cells C# | How to avoid custom IFilePathProvider in Aspose.Cells HTML export | Preserve worksheet sequence when saving as HTML Aspose | Export hidden sheets to HTML with Aspose.Cells
// Developer Intent: Save the full workbook as HTML while preserving the native sheet order using the default file‑path mapping.
// Use Cases: Generate separate HTML pages for each worksheet in a multi‑sheet workbook for web preview, maintaining the original order. | Create static HTML documentation that includes hidden worksheets without modifying file‑path settings. | Provide a quick, order‑preserving HTML export for client‑side rendering of Excel reports.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, preserving sheet order and using the default IFilePathProvider. | Explain why a custom IFilePathProvider can alter worksheet order in HTML export and how to keep the default behavior. | Show how to configure HtmlSaveOptions to include hidden worksheets and export all sheets while retaining the workbook's sheet sequence.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates exporting a workbook to HTML while keeping the original worksheet order.
    // No custom IFilePathProvider is set, so the default mapping is used and the order is preserved.
    // Demonstrates how to save an Aspose.Cells workbook as HTML while keeping the original worksheet sequence. The sample creates three sheets, uses HtmlSaveOptions with ExportActiveWorksheetOnly = false and ExportHiddenWorksheet = true, and relies on the default IFilePathProvider so the generated HTML files follow the workbook's sheet order.
    public class ExportWorkbookToHtml
    {
        public static void Main()
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Add sample worksheets and data.
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            sheet1.Cells["A1"].PutValue("Data in first sheet");

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            sheet2.Cells["A1"].PutValue("Data in second sheet");

            Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
            sheet3.Cells["A1"].PutValue("Data in third sheet");

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Keep default settings; do NOT assign a custom FilePathProvider.
                // This ensures the original worksheet order is retained.
                ExportActiveWorksheetOnly = false, // Export the whole workbook.
                ExportHiddenWorksheet = true        // Export hidden sheets if any.
            };

            // Save the workbook as HTML. The output will contain separate HTML files for each sheet,
            // and the order of the sheets will match the original workbook order.
            workbook.Save("WorkbookExport.html", saveOptions);

            Console.WriteLine("Workbook successfully exported to HTML with original sheet order preserved.");
        }
    }
}
