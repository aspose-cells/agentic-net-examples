// Title: Create an Excel workbook and add a column chart with sample data using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to create a new workbook, fill cells A1:B6 with category and value data, and insert a column chart spanning rows 5‑25 and columns 2‑11. | Demonstrate how to set the data range and assign a title to a column chart in an Aspose.Cells worksheet. | Explain the steps to save the workbook as an .xlsx file after the chart has been added with Aspose.Cells.
// Common Searches: Aspose.Cells .NET example for adding a column chart to a worksheet | C# code to set chart data range and title in Aspose.Cells | How to save an Excel file with a chart using Aspose.Cells in C# | Create sample data and column chart programmatically with Aspose.Cells | Aspose.Cells chart positioning rows columns C#
// Tags: aspocells create column chart c# | aspocells set chart data range | aspocells add worksheet sample data | aspocells save workbook as xlsx | aspocells chart positioning rows columns | aspocells column chart title

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, populates cells A1:B6 with category and numeric values, adds a column chart positioned from row 5, column 2 to row 25, column 11, sets the chart's data range to A1:B6, assigns the title "Sample Column Chart", and saves the file as ColumnChartOutput.xlsx.
public class CreateWorkbookWithColumnChart
{
    public static void Main()
    {
        // Initialize a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 5; i++)
        {
            sheet.Cells[$"A{i + 1}"].PutValue($"Item {i}");
            sheet.Cells[$"B{i + 1}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet (rows 5‑25, columns 2‑11)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 2, 25, 11);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart (including headers)
        chart.SetChartDataRange("A1:B6", true);

        // Set a title for the chart
        chart.Title.Text = "Sample Column Chart";

        // Save the workbook to a file
        workbook.Save("ColumnChartOutput.xlsx", SaveFormat.Xlsx);
    }
}
