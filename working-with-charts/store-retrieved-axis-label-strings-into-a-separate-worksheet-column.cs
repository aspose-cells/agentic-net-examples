// Title: Extract category axis labels from an Aspose.Cells chart and write them to a worksheet column in C#
// AI Prompts: Write C# code using Aspose.Cells to calculate a chart, call GetAxisTexts on the CategoryAxis, and store the returned strings into a designated column of the same worksheet. | Show how to iterate over the array returned by chart.CategoryAxis.GetAxisTexts() and populate cells in column D starting at row 2 with those labels.
// Common Searches: C# Aspose.Cells retrieve category axis text after chart calculation | store chart axis labels into worksheet cells using Aspose.Cells .NET | Aspose.Cells GetAxisTexts example for a column chart | write category axis values to column D in Excel with Aspose.Cells | extract chart axis strings and save to a separate column in C#
// Tags: chart category axis GetAxisTexts Aspose.Cells | write axis labels to worksheet column C# | calculate chart before extracting axis texts | populate Excel column with chart labels Aspose.Cells | Aspose.Cells chart data extraction to cells

using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, calculates the chart to generate axis texts, retrieves the category axis labels via GetAxisTexts, writes those labels into column D starting at row 2, and saves the file as AxisLabelsStored.xlsx.
class StoreAxisLabels
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
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the series and categories
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Calculate the chart to generate axis texts
        chart.Calculate();

        // Retrieve the category axis labels using the recommended GetAxisTexts method
        string[] axisLabels = chart.CategoryAxis.GetAxisTexts();

        // Store the retrieved labels into column D, starting from row 2
        int startRow = 1; // zero‑based index (row 2)
        int targetColumn = 3; // zero‑based index for column D
        for (int i = 0; i < axisLabels.Length; i++)
        {
            worksheet.Cells[startRow + i, targetColumn].PutValue(axisLabels[i]);
        }

        // Save the workbook
        workbook.Save("AxisLabelsStored.xlsx");
    }
}
