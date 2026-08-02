// Title: Aspose.Cells .NET Example: Create a Combo Chart with Column, Line (Secondary Axis) and Area Series in C#
// Description: C# code that demonstrates how to use Aspose.Cells for .NET to build an Excel workbook, fill it with month categories and three data series, add a combo chart that mixes a column series, a line series on a secondary axis, and an area series, set custom axis titles, and save the file as ComboChart.xlsx.
// Keywords: Aspose.Cells C# combo chart | combo chart secondary axis Aspose.Cells | column line area chart .NET | Excel chart multiple series Aspose | Aspose.Cells sample code | GitHub Aspose.Cells example | chart with different axes C#
// Common Searches: Aspose.Cells create combo chart with secondary axis | C# example for column, line and area series in one chart | How to add multiple chart types to the same Excel chart using Aspose.Cells | Set axis titles for combo chart Aspose.Cells .NET | Export combo chart to Excel with Aspose.Cells
// Developer Intent: Generate an Excel file that contains a combo chart combining column, line (on a secondary axis), and area series using Aspose.Cells for .NET.
// Use Cases: Display monthly sales as columns while showing profit margin as a line on a secondary scale and cumulative totals as an area chart. | Create a single‑page dashboard where each KPI uses the most suitable chart type and axis for clear comparison. | Export a combined visualization to Excel for stakeholders who need both primary and secondary value axes.
// AI Prompts: Show how to move the area series to the secondary axis instead of the line series in the Aspose.Cells combo chart example. | Provide code to add data labels to all three series of the combo chart created above. | Explain how to change the column series to a stacked column while keeping the line and area series unchanged.

using Aspose.Cells;
using Aspose.Cells.Charts;

// C# code that demonstrates how to use Aspose.Cells for .NET to build an Excel workbook, fill it with month categories and three data series, add a combo chart that mixes a column series, a line series on a secondary axis, and an area series, set custom axis titles, and save the file as ComboChart.xlsx.
class ComboChartExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Column");
        sheet.Cells["C1"].PutValue("Line");
        sheet.Cells["D1"].PutValue("Area");

        // Sample data
        string[] categories = { "Jan", "Feb", "Mar", "Apr" };
        double[] columnVals = { 10, 20, 30, 40 };
        double[] lineVals   = { 15, 25, 35, 45 };
        double[] areaVals   = { 5, 15, 25, 35 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(categories[i]);   // Category
            sheet.Cells[i + 1, 1].PutValue(columnVals[i]); // Column series values
            sheet.Cells[i + 1, 2].PutValue(lineVals[i]);   // Line series values
            sheet.Cells[i + 1, 3].PutValue(areaVals[i]);   // Area series values
        }

        // Add a combo chart (initially a Column chart)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];
        chart.Title.Text = "Combo Chart";

        // Series 1: Column (primary value axis)
        chart.NSeries.Add("=Sheet1!$B$2:$B$5", true);
        chart.NSeries[0].Type = ChartType.Column;
        chart.NSeries[0].Name = "Column Series";

        // Series 2: Line (secondary value axis)
        chart.NSeries.Add("=Sheet1!$C$2:$C$5", true);
        chart.NSeries[1].Type = ChartType.Line;
        chart.NSeries[1].Name = "Line Series";
        chart.NSeries[1].PlotOnSecondAxis = true; // Use secondary axis

        // Series 3: Area (primary value axis)
        chart.NSeries.Add("=Sheet1!$D$2:$D$5", true);
        chart.NSeries[2].Type = ChartType.Area;
        chart.NSeries[2].Name = "Area Series";

        // Set category (X) axis data
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";

        // Optional: axis titles for clarity
        chart.CategoryAxis.Title.Text = "Month";
        chart.ValueAxis.Title.Text = "Primary Value";
        chart.SecondValueAxis.Title.Text = "Secondary Value";

        // Save the workbook
        workbook.Save("ComboChart.xlsx");
    }
}
