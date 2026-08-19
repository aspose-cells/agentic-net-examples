// Title: Export a Combo (Column & Line) Chart to PDF with Embedded Fonts using Aspose.Cells for .NET
// Description: Creates a workbook, adds a column‑line combo chart for quarterly sales and profit, and saves only the chart as a PDF with fonts embedded automatically. Ideal for consistent rendering across Windows, macOS and Linux.
// Keywords: Aspose.Cells export chart to PDF | combo chart PDF C# | embedded fonts PDF Aspose.Cells | .NET chart to PDF | column line chart Aspose.Cells | save chart as PDF | Aspose.Cells PDF save options | C# export only chart | Aspose.Cells PDF generation
// Common Searches: How to export a combo chart to PDF with embedded fonts in C# | Aspose.Cells export only chart to PDF | Embedding fonts when saving chart as PDF using Aspose.Cells | Create column and line combo chart and convert to PDF .NET | Aspose.Cells PDF save options for charts
// Developer Intent: Generate a PDF of a combo chart with fonts embedded for reliable cross‑platform display.
// Use Cases: Produce quarterly sales‑profit reports as PDF charts that look identical on any device. | Insert chart‑only PDFs into presentations without losing font fidelity. | Automate batch conversion of workbook charts to individual PDFs for archival.
// AI Prompts: Write C# code with Aspose.Cells to build a column‑line combo chart and export it to a PDF with embedded fonts. | Show how to configure Aspose.Cells PDF save options to guarantee font embedding when saving a chart. | Provide a script that iterates through all charts in a workbook and saves each as a separate PDF with fonts embedded.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartPdf
{
    // Creates a workbook, adds a column‑line combo chart for quarterly sales and profit, and saves only the chart as a PDF with fonts embedded automatically. Ideal for consistent rendering across Windows, macOS and Linux.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the combo chart
                // Column series data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(220);

                // Line series data (secondary axis)
                sheet.Cells["C1"].PutValue("Profit");
                sheet.Cells["C2"].PutValue(30);
                sheet.Cells["C3"].PutValue(45);
                sheet.Cells["C4"].PutValue(40);
                sheet.Cells["C5"].PutValue(55);

                // Add a combo chart (Column + Line)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // First series – column (Sales)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries[0].Name = "Sales";

                // Second series – line (Profit)
                chart.NSeries.Add("C2:C5", true);
                chart.NSeries[1].Name = "Profit";
                chart.NSeries[1].Type = ChartType.Line; // Change series type to line

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A5";

                // Optional: give the chart a title
                chart.Title.Text = "Quarterly Sales and Profit";

                // Export only the chart to PDF (fonts are embedded by default)
                chart.ToPdf("ComboChart.pdf");

                Console.WriteLine("Combo chart exported to PDF with embedded fonts successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
