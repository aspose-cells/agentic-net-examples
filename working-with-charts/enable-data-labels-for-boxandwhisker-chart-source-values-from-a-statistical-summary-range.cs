using Aspose.Cells;
using Aspose.Cells.Charts;

class BoxWhiskerDataLabelsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate raw data for the box‑and‑whisker chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q1");
        sheet.Cells["A4"].PutValue("Q1");
        sheet.Cells["A5"].PutValue("Q2");
        sheet.Cells["A6"].PutValue("Q2");
        sheet.Cells["A7"].PutValue("Q2");

        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(15);
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["B5"].PutValue(20);
        sheet.Cells["B6"].PutValue(35);
        sheet.Cells["B7"].PutValue(40);

        // Statistical summary range (e.g., mean values) that will be used for data labels
        sheet.Cells["C1"].PutValue("Mean");
        sheet.Cells["C2"].PutValue(20);
        sheet.Cells["C3"].PutValue(30);
        sheet.Cells["C4"].PutValue(35);
        sheet.Cells["C5"].PutValue(25);
        sheet.Cells["C6"].PutValue(40);
        sheet.Cells["C7"].PutValue(45);

        // Add a box‑and‑whisker chart
        int chartIdx = sheet.Charts.Add(ChartType.BoxWhisker, 5, 0, 25, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Define the series data (values) and categories
        chart.NSeries.Add("B2:B7", true);
        chart.NSeries.CategoryData = "A2:A7";

        // Enable data labels and bind them to the statistical summary range
        Series series = chart.NSeries[0];
        series.DataLabels.ShowValue = true;          // Show the original value
        series.DataLabels.ShowCellRange = true;      // Use a cell range for label text
        series.DataLabels.LinkedSource = "C2:C7";    // Summary values to display
        series.DataLabels.Position = LabelPositionType.InsideBase; // Optional positioning

        // Save the workbook
        workbook.Save("BoxWhiskerDataLabels.xlsx");
    }
}