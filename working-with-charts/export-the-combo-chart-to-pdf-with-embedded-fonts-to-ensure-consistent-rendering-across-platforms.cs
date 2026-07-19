// Title: Export a Combo (Column‑Line) Chart to PDF with Embedded Fonts using Aspose.Cells for .NET
// Description: C# code that creates a column‑line combo chart, assigns month, sales, and profit data, and saves the chart as a PDF with fonts embedded automatically, ensuring the visual layout stays identical on every platform.
// Keywords: Aspose.Cells export chart to PDF | combo chart PDF Aspose | embed fonts PDF Aspose.Cells | C# Aspose.Cells Chart.ToPdf | Aspose.Cells PDF font embedding | export column line chart .NET | Aspose.Cells chart rendering consistency | PDF generation with embedded fonts | .NET chart to PDF example
// Common Searches: How to export a combo chart to PDF with embedded fonts using Aspose.Cells | Aspose.Cells C# save chart as PDF with font embedding | Export Aspose.Cells chart to PDF preserving fonts | Chart.ToPdf embed fonts Aspose.Cells | Create PDF from Excel chart with embedded fonts .NET
// Developer Intent: Generate a PDF of a combo chart while embedding all fonts for consistent rendering.
// Use Cases: Produce sales‑profit reports where the chart must look identical on client devices. | Automate PDF generation for dashboards that include mixed‑type charts. | Create distribution‑ready documents that avoid missing‑font issues on recipient machines.
// AI Prompts: Write C# code that builds a column‑line combo chart with Aspose.Cells and exports it to a PDF with embedded fonts. | Explain how Aspose.Cells embeds fonts when using Chart.ToPdf and how to verify the embedding in the output file. | Show how to style a combo chart (colors, markers, axis titles) before exporting it to PDF while keeping fonts embedded.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that creates a column‑line combo chart, assigns month, sales, and profit data, and saves the chart as a PDF with fonts embedded automatically, ensuring the visual layout stays identical on every platform.
class ExportComboChartPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the combo chart
            worksheet.Cells["A1"].PutValue("Month");
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["A4"].PutValue("Mar");

            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(120);
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["B4"].PutValue(180);

            worksheet.Cells["C1"].PutValue("Profit");
            worksheet.Cells["C2"].PutValue(30);
            worksheet.Cells["C3"].PutValue(45);
            worksheet.Cells["C4"].PutValue(60);

            // Add a combo chart (Column + Line) to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // First series – column type (Sales)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].Name = "Sales";

            // Second series – line type (Profit)
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Type = ChartType.Line; // Set series type to Line
            // Note: IsOnSecondaryAxis property is not available in this version; the series will share the primary axis.
            chart.NSeries[1].Name = "Profit";

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // Export the chart to PDF (fonts are embedded by default)
            chart.ToPdf("ComboChart.pdf");

            Console.WriteLine("Combo chart exported to PDF with embedded fonts successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
