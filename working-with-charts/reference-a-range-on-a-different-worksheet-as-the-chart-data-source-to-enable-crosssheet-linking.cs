// Title: Create a column chart in Aspose.Cells .NET with its data range taken from another worksheet
// AI Prompts: Write C# code that adds a column chart on a new worksheet and sets its data source to a range on a different worksheet using Aspose.Cells. | Show how to bind a chart title to a cell located on a separate sheet in an Aspose.Cells workbook.
// Common Searches: Aspose.Cells how to set chart data source to a range on a different sheet in C# | C# Aspose.Cells cross‑sheet chart data range example | link chart title to a cell on another worksheet using Aspose.Cells .NET | use A1 notation to reference external worksheet range for chart in Aspose.Cells | create column chart on separate worksheet with data from DataSheet Aspose.Cells
// Tags: set chart data range cross‑sheet Aspose.Cells | column chart on separate worksheet C# | chart title linked source Aspose.Cells | A1 notation range for chart data Aspose.Cells | reference another worksheet for chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds a data sheet and a chart sheet, fills sample data, inserts a column chart on the chart sheet, sets the chart's data range to a range on the data sheet using A1 notation, optionally links the chart title to a cell on the data sheet, and saves the file as CrossSheetChart.xlsx.
class CrossSheetChartExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // First worksheet will hold the data
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "DataSheet";

        // Second worksheet will contain the chart
        Worksheet chartSheet = workbook.Worksheets.Add("ChartSheet");

        // Populate sample data in DataSheet
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["A2"].PutValue("A");
        dataSheet.Cells["B2"].PutValue(10);
        dataSheet.Cells["A3"].PutValue("B");
        dataSheet.Cells["B3"].PutValue(20);
        dataSheet.Cells["A4"].PutValue("C");
        dataSheet.Cells["B4"].PutValue(30);

        // Add a column chart to ChartSheet
        int chartIndex = chartSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = chartSheet.Charts[chartIndex];

        // Set the chart's data range to a range on a different worksheet (cross‑sheet linking)
        // The range string follows Excel A1 notation and includes the sheet name.
        chart.SetChartDataRange("'DataSheet'!$A$1:$B$4", true);

        // Optionally link the chart title to a cell on the data sheet
        dataSheet.Cells["D1"].PutValue("Sales Chart");
        chart.Title.LinkedSource = "'DataSheet'!$D$1";

        // Save the workbook
        workbook.Save("CrossSheetChart.xlsx");
    }
}
