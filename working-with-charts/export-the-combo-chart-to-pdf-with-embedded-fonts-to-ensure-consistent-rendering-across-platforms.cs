// Title: Export a Combo Column‑Line Chart to PDF with Embedded Fonts using Aspose.Cells for .NET (C#)
// Description: Creates a workbook with quarterly sales and profit data, builds a combo chart (column for sales, line for profit), sets the category axis and title, then uses Chart.ToPdf to generate a PDF where fonts are embedded automatically, guaranteeing identical appearance on any device.
// Keywords: Aspose.Cells chart to PDF | combo chart PDF export | embed fonts PDF Aspose | C# Aspose.Cells chart export | column line chart PDF | Chart.ToPdf embedded fonts
// Common Searches: aspocells export combo chart pdf | how to embed fonts in PDF chart aspnet | c# chart topdf embed fonts | save excel combo chart as pdf aspose | pdf font embedding aspose.cells chart
// Developer Intent: Produce a PDF of a combo column‑line chart with fonts embedded for reliable cross‑platform rendering.
// Use Cases: Generate printable quarterly sales reports with consistent chart appearance. | Automate PDF creation for dashboards that require exact font matching across browsers. | Batch export multiple Excel worksheets' combo charts to PDF archives while preserving typography.
// AI Prompts: Write C# code that creates a combo column‑line chart from worksheet data and saves it as a PDF with embedded fonts using Aspose.Cells. | Explain the default font‑embedding behavior of Chart.ToPdf and how to verify embedded fonts in the output PDF. | Show how to configure EmbedStandardWindowsFonts property before exporting a chart to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartPdf
{
    // Creates a workbook with quarterly sales and profit data, builds a combo chart (column for sales, line for profit), sets the category axis and title, then uses Chart.ToPdf to generate a PDF where fonts are embedded automatically, guaranteeing identical appearance on any device.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);
            sheet.Cells["C5"].PutValue(80);

            // Add a combo chart (column + line) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // First series – column (Sales)
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Sales";

            // Second series – line (Profit)
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].Name = "Profit";
            chart.NSeries[1].Type = ChartType.Line; // set second series as line to create a combo chart

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Optional: set a title
            chart.Title.Text = "Quarterly Sales and Profit";

            // Export the chart to PDF.
            // The Chart.ToPdf method embeds fonts by default (EmbedStandardWindowsFonts = true).
            chart.ToPdf("ComboChart.pdf");

            Console.WriteLine("Combo chart exported to PDF with embedded fonts successfully.");
        }
    }
}
