using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with long category names to demonstrate wrapping
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
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Access the data labels of the first series
        DataLabels dataLabels = chart.NSeries[0].DataLabels;
        dataLabels.ShowValue = true;               // Show the numeric values
        dataLabels.Position = LabelPositionType.Center;

        // Enable text wrapping so that long category names become multi‑line
        dataLabels.IsTextWrapped = true;

        // Allow the label shape to automatically resize to fit the wrapped text
        dataLabels.IsResizeShapeToFitText = true;

        // Set a narrow width to force wrapping; the shape will expand vertically
        dataLabels.WidthPixel = 80;

        // Save the workbook
        workbook.Save("DataLabelsWrappedResized.xlsx");
    }
}