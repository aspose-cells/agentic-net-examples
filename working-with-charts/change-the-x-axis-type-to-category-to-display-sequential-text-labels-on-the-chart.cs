// Title: Aspose.Cells C# – Set X‑Axis to Category Scale for Text Labels in a Column Chart
// Description: Creates a workbook, fills cells with string labels and numeric values, adds a column chart, binds the series and category data, switches the X‑axis to a CategoryScale so the labels are displayed sequentially, and saves the result as an XLSX file.
// Keywords: Aspose.Cells | C# chart example | CategoryScale | X axis category | column chart | text labels | Excel automation | chart axis type | programmatic Excel | Aspose.Cells tutorial
// Common Searches: Aspose.Cells change X axis to category scale C# | display string labels on chart axis Aspose.Cells | set CategoryAxis.CategoryType in .NET | column chart with custom X‑axis labels using Aspose.Cells | how to use CategoryScale for chart axis in C#
// Developer Intent: Configure the chart’s X‑axis as a category axis so that sequential text labels are shown instead of numeric values.
// Use Cases: Generate a sales dashboard where month names appear as X‑axis labels in a column chart. | Create a product performance report with custom category names on the chart’s horizontal axis. | Export an Excel workbook that visualizes survey results using non‑numeric X‑axis labels.
// AI Prompts: Write C# code with Aspose.Cells that sets the X‑axis of a line chart to CategoryScale and binds it to a range of string cells. | Explain the steps to convert a chart’s X‑axis from Value to Category type and update its data source in Aspose.Cells. | Show how to programmatically verify that CategoryAxis.CategoryType equals CategoryScale before saving the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills cells with string labels and numeric values, adds a column chart, binds the series and category data, switches the X‑axis to a CategoryScale so the labels are displayed sequentially, and saves the result as an XLSX file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data with sequential text labels for the X axis
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Label " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range and category (X‑axis) data
        chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);
        chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

        // Change the X axis type to CategoryScale to display the text labels sequentially
        chart.CategoryAxis.CategoryType = CategoryType.CategoryScale;

        // Save the workbook
        workbook.Save("ChartWithCategoryAxis.xlsx", SaveFormat.Xlsx);
    }
}
