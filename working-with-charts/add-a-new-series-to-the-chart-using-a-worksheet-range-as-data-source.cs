// Title: Add a New Series to an Aspose.Cells Column Chart from a Worksheet Range (C#)
// Description: Creates a workbook, fills A1:C4 with categories and two data series, inserts a column chart, adds the first series from B2:B4 with categories from A2:A4, then adds a second series from C2:C4 and sets its name from cell C1, finally saves the file as ChartWithMultipleSeries.xlsx.
// Keywords: Aspose.Cells chart series C# | NSeries.Add range | add series to column chart Aspose | set series name from cell Aspose.Cells | multiple series chart .NET
// Common Searches: how to add another series to an Aspose.Cells chart | Aspose.Cells set series name from header cell | add multiple data series to column chart C# | use NSeries.Add with worksheet range Aspose
// Developer Intent: Programmatically add an additional data series to an existing Aspose.Cells chart by referencing a worksheet range and optionally naming the series from a header cell.
// Use Cases: Compare monthly sales of two products in a single column chart, each product sourced from its own column range. | Build a financial dashboard where revenue and expense columns are added as separate series to a chart automatically. | Allow users to select extra data columns in a spreadsheet and dynamically extend the chart with new series and proper labels.
// AI Prompts: Generate C# code that adds three series to a line chart in Aspose.Cells, each using a different column range and naming them from the header cells. | Explain how to update the data range of an existing chart series after modifying worksheet values in Aspose.Cells. | Provide step‑by‑step instructions to add a series to a chart, set its category axis, and assign a series name using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills A1:C4 with categories and two data series, inserts a column chart, adds the first series from B2:B4 with categories from A2:A4, then adds a second series from C2:C4 and sets its name from cell C1, finally saves the file as ChartWithMultipleSeries.xlsx.
class AddSeriesExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");

        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Add the first series using a worksheet range as data source
        chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);
        // Set the category (X‑axis) data for the series
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

        // Add a new series to the same chart using another range
        chart.NSeries.Add("=Sheet1!$C$2:$C$4", true);
        // Optionally set the name of the newly added series from cell C1
        chart.NSeries.SetSeriesNames(1, "C1", true);

        // Save the workbook with the chart
        workbook.Save("ChartWithMultipleSeries.xlsx");
    }
}
