// Title: Export a column chart to a Letter-size PDF with top‑left alignment using Aspose.Cells in C#
// AI Prompts: Generate C# code that creates a column chart from a worksheet range and uses Aspose.Cells.Chart.ToPdf to place the chart at the top‑left corner of an 8.5×11 inches PDF page. | Update an existing Aspose.Cells chart export so that the ToPdf call includes PageLayoutAlignmentType.Left and PageLayoutAlignmentType.Top for precise positioning.
// Common Searches: how to position an Aspose.Cells chart at the top left of a PDF page in C# | Aspose.Cells ToPdf method alignment options for charts | exporting Excel chart to letter size PDF with specific alignment using Aspose.Cells | C# Aspose.Cells chart PDF page layout left top alignment example | set chart placement when converting to PDF with Aspose.Cells
// Tags: Aspose.Cells chart ToPdf alignment | C# export chart to PDF top left | Aspose.Cells page layout alignment types | column chart PDF positioning Aspose.Cells | letter size PDF chart export C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAlignmentDemo
{
    // The example creates a workbook, fills it with sample data, adds a column chart, defines its data range, and then exports the chart to a PDF named ChartTopLeft.pdf. The ToPdf method is called with page dimensions of 8.5 × 11 inches and PageLayoutAlignmentType.Left/Top to position the chart at the top‑left corner of the page.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);
            chart.Title.Text = "Sample Chart";

            // Export the chart to PDF, aligning it to the top‑left corner of the page
            // Page size: 8.5 x 11 inches (standard Letter)
            // Horizontal alignment: Left, Vertical alignment: Top
            chart.ToPdf("ChartTopLeft.pdf", 8.5f, 11f,
                        PageLayoutAlignmentType.Left,
                        PageLayoutAlignmentType.Top);

            Console.WriteLine("Chart exported to PDF with top‑left alignment.");
        }
    }
}
