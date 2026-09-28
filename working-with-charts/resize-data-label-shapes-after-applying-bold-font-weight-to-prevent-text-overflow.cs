// Title: How to resize chart data label shapes after applying bold font in Aspose.Cells for .NET to avoid overflow
// AI Prompts: Create a column chart, enable data labels, set the label font to bold, then turn off automatic shape resizing and specify WidthPixel and HeightPixel values using Aspose.Cells. | Write C# code that adds a chart, makes data label text bold, and manually sets the label shape dimensions to prevent text clipping. | Demonstrate how to disable auto‑resize for chart data labels and assign custom pixel width and height after applying a bold font with Aspose.Cells.
// Common Searches: Aspose.Cells C# set fixed size for chart data labels after making font bold | prevent data label text overflow in column chart using Aspose.Cells .NET | how to disable automatic resizing of data label shapes in Aspose.Cells chart | custom width and height for chart data labels Aspose.Cells example | adjust data label shape dimensions after applying bold font in Aspose.Cells
// Tags: set data label shape size Aspose.Cells | disable data label auto resize .NET | bold font data labels chart Aspose.Cells | custom label dimensions WidthPixel HeightPixel | column chart data label overflow fix

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, adds a column chart, enables data labels, applies a bold font to the labels, disables automatic shape resizing, sets the label width to 80 pixels and height to 30 pixels, and saves the file as ResizedDataLabels.xlsx.
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

        // Add a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = worksheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable data labels for the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;

        // Apply bold font to the data labels
        dataLabels.Font.IsBold = true;
        // Propagate the font change to all child label nodes
        dataLabels.ApplyFont();

        // Disable automatic shape resizing and set a custom size
        // This prevents the bold text from overflowing the label shape
        dataLabels.IsResizeShapeToFitText = false;
        dataLabels.WidthPixel = 80;   // Adjust width as needed
        dataLabels.HeightPixel = 30;  // Adjust height as needed

        // Save the workbook
        workbook.Save("ResizedDataLabels.xlsx");
    }
}
