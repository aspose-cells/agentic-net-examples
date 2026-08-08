// Title: Export Merged Excel Workbooks with Charts and Shapes to PDF using Aspose.Cells for .NET (C#)
// Description: Creates a workbook with a column chart and another with a rectangle shape, merges the second into the first via Workbook.Combine, ensures the output folder exists, and saves the combined file as a PDF to confirm that charts and drawing objects retain their layout.
// Keywords: Aspose.Cells | C# | Workbook.Combine | export to PDF | Excel chart PDF | shape PDF rendering | merge Excel files | .NET PDF conversion | visual fidelity PDF | combined workbook PDF
// Common Searches: Aspose.Cells merge workbooks and export to PDF | C# export Excel chart and shape to PDF | How to combine two Excel files and save as PDF using Aspose | Save workbook with drawings as PDF .NET | Verify chart rendering in PDF with Aspose.Cells
// Developer Intent: Combine several Excel workbooks and generate a single PDF that preserves charts and drawing objects.
// Use Cases: Produce a consolidated PDF report that includes charts from one source workbook and graphic placeholders from another. | Automate creation of printable portfolios by merging data sheets with visual elements before PDF conversion. | Run quality‑assurance checks on merged workbooks by exporting them to PDF and reviewing layout consistency.
// AI Prompts: Generate C# code that merges three workbooks, each containing a different chart type, and exports the result to PDF while keeping all formatting. | Show how to add robust error handling for missing data or invalid chart ranges when saving a combined workbook to PDF with Aspose.Cells. | Explain how to configure PdfSaveOptions to improve image and shape quality after merging workbooks in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// Creates a workbook with a column chart and another with a rectangle shape, merges the second into the first via Workbook.Combine, ensures the output folder exists, and saves the combined file as a PDF to confirm that charts and drawing objects retain their layout.
class ExportCombinedWorkbookToPdf
{
    static void Main()
    {
        try
        {
            // ---------- Create first workbook with a sample chart ----------
            Workbook wb1 = new Workbook(); // create workbook
            Worksheet ws1 = wb1.Worksheets[0];
            ws1.Name = "Data1";

            // Populate data for the chart
            ws1.Cells["A1"].PutValue("Category");
            ws1.Cells["A2"].PutValue("Apple");
            ws1.Cells["A3"].PutValue("Banana");
            ws1.Cells["B1"].PutValue("Value");
            ws1.Cells["B2"].PutValue(30);
            ws1.Cells["B3"].PutValue(45);

            // Add a column chart and bind it to the data range
            int chartIdx1 = ws1.Charts.Add(ChartType.Column, 5, 0, 15, 10);
            Chart chart1 = ws1.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B3", true);
            chart1.NSeries.CategoryData = "A2:A3";
            chart1.Title.Text = "Fruit Chart";

            // ---------- Create second workbook with a sample shape (as image placeholder) ----------
            Workbook wb2 = new Workbook(); // create second workbook
            Worksheet ws2 = wb2.Worksheets[0];
            ws2.Name = "Data2";

            ws2.Cells["A1"].PutValue("Sample Text");

            // Add a rectangle shape to act as an image placeholder
            // Parameters: type, upperLeftRow, upperLeftColumn, top, left, height, width
            Shape shape = ws2.Shapes.AddShape(MsoDrawingType.Rectangle, 5, 0, 5, 0, 200, 100);
            shape.Placement = PlacementType.FreeFloating;
            shape.Text = "Image Placeholder";

            // ---------- Combine the two workbooks ----------
            wb1.Combine(wb2); // combine wb2 into wb1

            // ---------- Export the combined workbook to PDF ----------
            string outputPath = "CombinedWorkbook.pdf";

            // Ensure the directory exists (in case a relative path is used)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb1.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine("Combined workbook exported to PDF successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
