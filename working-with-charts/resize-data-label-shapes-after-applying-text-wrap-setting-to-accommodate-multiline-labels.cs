// Title: Resize chart data label shapes to fit wrapped text in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# using Aspose.Cells to turn on text wrapping for chart data labels and trigger automatic shape resizing so the labels grow to fit multiple lines. | Specify a starting pixel width for data label shapes while permitting height growth after text wrapping in a column chart.
// Common Searches: Aspose.Cells C# enable text wrap on chart data labels and auto‑size shape | how to set a pixel width for data label shapes in an Aspose.Cells column chart | make data labels expand vertically for long category names using Aspose.Cells | C# Aspose.Cells chart data label shape resizing after wrapping text | adjust chart data label dimensions to fit wrapped text in .NET
// Tags: chart data label shape resize based on wrapped text Aspose.Cells | enable text wrap for chart data labels .NET | initial label width setting Aspose.Cells | column chart multi‑line category labels C# | Aspose.Cells automatic label shape adjustment

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a column chart with long category names, enables text wrapping on the series data labels, activates automatic shape resizing, sets an initial label width in pixels, and saves the file as ResizeDataLabelShapes.xlsx.
class ResizeDataLabelShapes
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with long category names to force wrapping
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Very Long Category Name 1");
        worksheet.Cells["A3"].PutValue("Very Long Category Name 2");
        worksheet.Cells["A4"].PutValue("Very Long Category Name 3");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(250);
        worksheet.Cells["B4"].PutValue(370);

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Access the data labels of the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;

        // Show the values on the data labels
        dataLabels.ShowValue = true;

        // Enable text wrapping so that long category names become multi‑line
        dataLabels.IsTextWrapped = true;

        // Enable auto‑resize of the label shape to fit the wrapped text
        dataLabels.IsResizeShapeToFitText = true;

        // Optional: set an initial width; the shape will expand vertically to accommodate wrapped lines
        dataLabels.WidthPixel = 80;

        // Save the workbook
        workbook.Save("ResizeDataLabelShapes.xlsx");
    }
}
