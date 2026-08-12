// Title: Export Merged Cells to HTML and Verify colspan/rowspan with Aspose.Cells for .NET
// Description: Creates a workbook, merges range A1:C2, saves it as HTML using Aspose.Cells, then checks the generated markup for colspan="3" and rowspan="2" to confirm that the merged layout is preserved.
// Keywords: Aspose.Cells HTML export | merged cells to HTML | colspan rowspan verification | C# Aspose.Cells HtmlSaveOptions | Excel to HTML conversion | preserve merged layout | automated HTML validation
// Common Searches: Aspose.Cells export merged cells to HTML | verify colspan and rowspan in Aspose.Cells HTML output | C# merge A1:C2 and save as HTML | how to keep merged cells when converting Excel to HTML | Aspose.Cells HtmlSaveOptions merged cell settings
// Developer Intent: Generate an HTML file from a workbook that contains merged cells and programmatically confirm that the merged structure appears correctly in the markup.
// Use Cases: Publish spreadsheet reports with header rows that span multiple columns on a website. | Create web‑ready Excel exports where merged titles must render with proper colspan/rowspan tags. | Automate regression tests that validate HTML output after Excel‑to‑HTML conversion.
// AI Prompts: Write C# code using Aspose.Cells to merge cells A1:C2, export the workbook to HTML, and assert that the HTML contains colspan="3" and rowspan="2". | Explain how HtmlSaveOptions handles merged cells during HTML conversion and how to adjust settings if the default behavior changes. | Generate a C# unit test that loads a workbook with merged cells, saves it as HTML, and verifies the presence of correct colspan and rowspan attributes.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMergedHtmlDemo
{
    // Creates a workbook, merges range A1:C2, saves it as HTML using Aspose.Cells, then checks the generated markup for colspan="3" and rowspan="2" to confirm that the merged layout is preserved.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Put a value in the top-left cell of the merged area
            cells["A1"].PutValue("Merged Header");

            // Merge cells A1:C2 (rows 0-1, columns 0-2)
            // firstRow = 0, firstColumn = 0, totalRows = 2, totalColumns = 3
            cells.Merge(0, 0, 2, 3);

            // Verify merged areas before saving
            CellArea[] mergedAreas = cells.GetMergedAreas();
            Console.WriteLine($"Number of merged areas: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine($"Merged area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                                  $"EndRow={area.EndRow}, EndColumn={area.EndColumn}");
            }

            // Configure HTML save options (default settings preserve merged cells)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Keep default behavior; can explicitly set if needed
                MergeEmptyTdType = MergeEmptyTdType.Default
            };

            // Save the workbook to HTML
            string htmlPath = "MergedCells.html";
            workbook.Save(htmlPath, htmlOptions);
            Console.WriteLine($"Workbook saved to HTML: {htmlPath}");

            // Simple verification: check the generated HTML for colspan/rowspan attributes
            string htmlContent = File.ReadAllText(htmlPath);
            bool hasColSpan = htmlContent.Contains("colspan=\"3\"");
            bool hasRowSpan = htmlContent.Contains("rowspan=\"2\"");

            Console.WriteLine("Verification of merged layout in HTML:");
            Console.WriteLine($"  colspan=\"3\" found: {hasColSpan}");
            Console.WriteLine($"  rowspan=\"2\" found: {hasRowSpan}");
        }
    }
}
