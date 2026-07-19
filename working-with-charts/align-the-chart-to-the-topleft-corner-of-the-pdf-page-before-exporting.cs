// Title: C# – Align an Aspose.Cells Chart to the Top‑Left Corner When Exporting to PDF
// Description: Creates a workbook, adds sample data and a column chart, then uses the Chart.ToPdf overload with page dimensions and PageLayoutAlignmentType.Left / Top to position the chart at the top‑left of an 8.5 × 11 in PDF page.
// Keywords: Aspose.Cells | C# chart to PDF | chart alignment | PageLayoutAlignmentType | top left PDF | export chart PDF | Aspose.Cells ToPdf | chart positioning | PDF page layout Aspose | Excel chart PDF export
// Common Searches: Aspose.Cells align chart top left PDF | chart.ToPdf alignment C# example | position chart in PDF using Aspose.Cells | set chart placement when exporting to PDF Aspose | Aspose.Cells PDF page layout options
// Developer Intent: Place a chart at the top‑left of a PDF page during export with Aspose.Cells in C#.
// Use Cases: Generate PDF reports where every chart starts at the page’s top‑left for a uniform layout. | Create invoices or dashboards that require a sales chart positioned at the top‑left to match branding. | Automate batch conversion of multiple worksheets, ensuring each exported chart aligns consistently at the top‑left of its PDF page.
// AI Prompts: Show C# code that exports an Aspose.Cells chart to a PDF and aligns it to the top‑left corner using the ToPdf method. | Explain how PageLayoutAlignmentType.Left and PageLayoutAlignmentType.Top affect chart placement in Aspose.Cells PDF export. | Provide a step‑by‑step guide for positioning a chart at a custom location (e.g., top‑left) when converting Excel to PDF with Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data and a column chart, then uses the Chart.ToPdf overload with page dimensions and PageLayoutAlignmentType.Left / Top to position the chart at the top‑left of an 8.5 × 11 in PDF page.
class AlignChartTopLeftPdf
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

        // Set the data source for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Sample Chart";

        // Export the chart to PDF, aligning it to the top‑left corner of the page
        string pdfFile = "ChartTopLeft.pdf";
        chart.ToPdf(pdfFile, 8.5f, 11f,
            PageLayoutAlignmentType.Left,   // Horizontal alignment
            PageLayoutAlignmentType.Top);   // Vertical alignment
    }
}
