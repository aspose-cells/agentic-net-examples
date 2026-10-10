// Title: Create a PDF bookmark that points to a chart by defining a named destination using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a workbook, adds a column chart, defines a named destination for the chart, and saves the workbook as a PDF with a bookmark linked to that destination using Aspose.Cells. | Show how to assign a custom name to a chart and use it as the target of a PDF bookmark in Aspose.Cells. | Write a script that ensures the output folder exists, then exports the workbook to PDF while preserving the chart bookmark.
// Common Searches: Aspose.Cells C# add PDF bookmark to a chart | link Excel chart to PDF bookmark with Aspose.Cells | export chart as PDF with clickable bookmark in .NET | C# create PDF destination for chart using Aspose.Cells | save workbook to PDF and generate chart bookmark automatically
// Tags: Aspose.Cells chart PDF link | C# set chart destination identifier | export workbook to PDF with chart reference | ensure output folder exists before PDF save Aspose.Cells | chart identifier for PDF navigation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsPdfBookmarkExample
{
    // The example creates a workbook, fills it with sample data, inserts a column chart, assigns a name to the chart, and saves the workbook as a PDF. A named destination is defined for the chart so that a PDF bookmark can navigate directly to the chart, and the code ensures the output directory exists before saving.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(15);

                // Insert a column chart into the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Optional: give the chart a name (useful for identification)
                chart.Name = "SalesChart";

                // Prepare PDF save options (bookmarks are enabled by default for worksheets)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Define output file path
                string outputPath = "ChartWithPdfBookmark.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
