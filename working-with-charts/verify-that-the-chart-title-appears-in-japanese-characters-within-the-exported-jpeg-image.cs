using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class VerifyJapaneseChartTitle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("項目");          // "Item" in Japanese
        sheet.Cells["A2"].PutValue("りんご");      // "Apple"
        sheet.Cells["A3"].PutValue("みかん");      // "Orange"
        sheet.Cells["A4"].PutValue("バナナ");      // "Banana"

        sheet.Cells["B1"].PutValue("数量");        // "Quantity"
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(80);
        sheet.Cells["B4"].PutValue(150);

        // Set a default font that supports Japanese characters
        // This ensures that the title is rendered correctly when the chart is saved as an image
        workbook.DefaultStyle.Font.Name = "MS Gothic";

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Set the chart title using Japanese characters
        chart.Title.Text = "日本語のタイトル";   // "Japanese Title"
        chart.Title.IsVisible = true;

        // Export the chart to a JPEG image
        // The overload with ImageOrPrintOptions allows us to specify the default font,
        // but if it is not available we rely on the workbook's default font set above.
        chart.ToImage("JapaneseChart.jpg", ImageType.Jpeg);

        // At this point, the file "JapaneseChart.jpg" contains the chart.
        // Verify manually that the title appears correctly in Japanese characters.
        Console.WriteLine("Chart exported to JPEG. Please open 'JapaneseChart.jpg' to verify the Japanese title.");
    }
}