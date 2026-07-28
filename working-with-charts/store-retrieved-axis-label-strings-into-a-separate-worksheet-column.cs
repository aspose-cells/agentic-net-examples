// Title: Store Chart Category Axis Labels in a Separate Worksheet Column with Aspose.Cells (C#)
// Description: Shows how to calculate a chart, retrieve its category axis texts via GetAxisTexts, and write each label into column A of a newly added worksheet, then save the workbook as an .xlsx file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | C# | chart axis labels | GetAxisTexts | chart.Calculate | write axis texts to worksheet | extract category axis | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells retrieve chart category axis labels | write chart axis texts to another sheet C# | GetAxisTexts example Aspose.Cells | store chart axis values in worksheet column | Aspose.Cells chart.Calculate usage
// Developer Intent: Extract a chart’s category axis labels and save them in a new worksheet column.
// Use Cases: Create a summary sheet that lists all categories displayed in a chart for reporting. | Generate a data‑validation list based on chart categories for downstream worksheets. | Export axis labels to a separate sheet for statistical analysis or external processing.
// AI Prompts: Provide C# code to extract both category and value axis labels from an Aspose.Cells chart and store them in separate columns. | Show how to loop through all charts in a workbook and save each chart’s axis texts to its own worksheet using Aspose.Cells. | Explain how to handle empty or null axis labels when writing them to a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Shows how to calculate a chart, retrieve its category axis texts via GetAxisTexts, and write each label into column A of a newly added worksheet, then save the workbook as an .xlsx file using Aspose.Cells for .NET.
class StoreAxisLabels
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        dataSheet.Cells["A1"].PutValue("Category");
        dataSheet.Cells["B1"].PutValue("Value");
        dataSheet.Cells["A2"].PutValue("A");
        dataSheet.Cells["B2"].PutValue(10);
        dataSheet.Cells["A3"].PutValue("B");
        dataSheet.Cells["B3"].PutValue(20);
        dataSheet.Cells["A4"].PutValue("C");
        dataSheet.Cells["B4"].PutValue(30);

        // Add a column chart linked to the data range
        int chartIndex = dataSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = dataSheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";     // Categories

        // Calculate the chart so that axis texts become available
        chart.Calculate();

        // Retrieve the category axis labels using the recommended GetAxisTexts method
        string[] axisLabels = chart.CategoryAxis.GetAxisTexts();

        // Create a new worksheet to store the retrieved labels
        Worksheet labelSheet = workbook.Worksheets.Add("AxisLabels");

        // Write each label into column A (index 0) starting from row 0 (cell A1)
        for (int i = 0; i < axisLabels.Length; i++)
        {
            labelSheet.Cells[i, 0].PutValue(axisLabels[i]);
        }

        // Save the workbook with the chart and the extracted axis labels
        workbook.Save("AxisLabelsOutput.xlsx");
    }
}
