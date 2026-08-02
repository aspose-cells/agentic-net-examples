// Title: Export a Waterfall Chart to PDF with Aspose.Cells for .NET (C#)
// Description: Loads "WaterfallChart.xlsx", accesses the first worksheet and its first chart, then uses Aspose.Cells' Chart.ToPdf method to render the Waterfall chart as an embedded image in "WaterfallChart.pdf".
// Keywords: Aspose.Cells | C# | Waterfall chart | Export chart to PDF | Chart.ToPdf | embed chart image | Excel to PDF conversion | .NET chart export
// Common Searches: Aspose.Cells export chart to PDF C# | How to save an Excel chart as PDF using Aspose | Waterfall chart PDF conversion Aspose.Cells | Chart.ToPdf example .NET | Convert specific chart to PDF programmatically
// Developer Intent: Convert a Waterfall chart inside an Excel workbook to a PDF file with the chart rendered as an image.
// Use Cases: Create a PDF report that contains only the Waterfall chart from a financial model. | Generate printable PDFs of selected charts for slide decks without exporting full worksheets. | Automate batch processing to save each chart in a workbook as an individual PDF for archiving.
// AI Prompts: Show a C# snippet that selects a chart by name and saves it as a PDF using Aspose.Cells. | Explain how to set page size, orientation, and margins when exporting a chart to PDF with Aspose.Cells. | Provide code that iterates through all charts in a workbook and creates separate PDF files for each.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads "WaterfallChart.xlsx", accesses the first worksheet and its first chart, then uses Aspose.Cells' Chart.ToPdf method to render the Waterfall chart as an embedded image in "WaterfallChart.pdf".
class WaterfallChartToPdf
{
    static void Main()
    {
        // Load the workbook that contains the Waterfall chart
        Workbook workbook = new Workbook("WaterfallChart.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the first chart in the worksheet (assumed to be the Waterfall chart)
        Chart chart = worksheet.Charts[0];

        // Export the chart directly to a PDF file.
        // The chart is rendered as an image inside the PDF.
        chart.ToPdf("WaterfallChart.pdf");

        // (Optional) Save the original workbook if you need a copy.
        // workbook.Save("WaterfallChart_copy.xlsx");
    }
}
