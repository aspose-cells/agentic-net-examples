// Title: C# – Export Excel Chart as PNG and Centered PDF Page with Localized “Other” Label using Aspose.Cells
// Description: Creates a workbook with a pie chart, sets the chart title to a localized "Other" label, exports the chart to a PNG image, and saves it as an 8.5×11 in PDF page centered on the page. Includes sample code for looping through multiple charts.
// Keywords: Aspose.Cells chart export | C# export chart to PDF | localize chart title Aspose.Cells | chart to PNG Aspose.Cells | center chart on PDF page | globalization Excel chart | Aspose.Cells PDF page layout
// Common Searches: How to export an Excel chart as a centered PDF page with Aspose.Cells C# | Apply a localized label to a chart title before exporting with Aspose.Cells | Extract chart image to PNG and convert to PDF using Aspose.Cells .NET | Loop through all charts in a workbook and export each to PDF Aspose.Cells
// Developer Intent: Export a chart image, apply a language‑specific label, and generate a centered PDF page for each chart using Aspose.Cells for .NET.
// Use Cases: Produce multilingual PDF reports where each chart occupies its own centered page. | Create PNG assets for web dashboards while also providing print‑ready PDF versions. | Automate batch processing of workbooks: translate chart titles and export each chart to separate PDF pages.
// AI Prompts: Generate C# code that iterates through all charts in a workbook, reads a translated title from a resource file, sets the title, and exports each chart to a centered PDF page of custom size using Aspose.Cells. | Show how to replace a chart's title with a localized string and then save the chart as both PNG and PDF with Aspose.Cells in .NET. | Explain how to control image quality and DPI when converting a chart to PNG before embedding it in a PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartExportExample
{
    // Creates a workbook with a pie chart, sets the chart title to a localized "Other" label, exports the chart to a PNG image, and saves it as an 8.5×11 in PDF page centered on the page. Includes sample code for looping through multiple charts.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Other"); // Category that will be localized
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["B3"].PutValue(35);
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a localized label "Other" to the chart title (or you could set a data label)
            chart.Title.Text = "Other";

            // Export the chart as an image (optional, demonstrates image extraction)
            string imagePath = "ChartImage.png";
            chart.ToImage(imagePath, ImageType.Png);
            Console.WriteLine($"Chart image saved to {imagePath}");

            // Export the chart to a PDF page with specific page size and centered alignment
            string pdfPath = "ChartPage.pdf";
            chart.ToPdf(pdfPath, 8.5f, 11f, PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);
            Console.WriteLine($"Chart PDF page saved to {pdfPath}");

            // If you have multiple charts, you can loop through them and export each one
            // Example loop (uncomment if needed):
            /*
            for (int i = 0; i < sheet.Charts.Count; i++)
            {
                Chart c = sheet.Charts[i];
                c.Title.Text = "Other";
                string imgFile = $"ChartImage_{i}.png";
                c.ToImage(imgFile, ImageType.Png);
                string pdfFile = $"ChartPage_{i}.pdf";
                c.ToPdf(pdfFile, 8.5f, 11f, PageLayoutAlignmentType.Center, PageLayoutAlignmentType.Center);
            }
            */

            // Save the workbook (optional, to keep the source Excel file)
            workbook.Save("SourceWorkbook.xlsx");
        }
    }
}
