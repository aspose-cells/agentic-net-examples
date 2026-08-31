// Title: Generate a column chart, update a data point, and save the workbook as ChartReport.xlsx using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a column chart from a data range, changes the value of a specific data point, and saves the workbook to ChartReport.xlsx. | Show how to programmatically modify the series data of an existing chart in a workbook and export the updated file as a new XLSX using Aspose.Cells.
// Common Searches: how to add a column chart to a new workbook and change a series value with Aspose.Cells C# | Aspose.Cells C# update chart data point and save as new Excel file | save modified Excel chart to a different filename using Aspose.Cells .NET | C# Aspose.Cells example for creating chart, editing data, and exporting to ChartReport.xlsx
// Tags: create column chart Aspose.Cells C# | modify chart series value Aspose.Cells | save workbook as xlsx Aspose.Cells | update chart data range Aspose.Cells .NET | export chart report Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// // This program creates a new workbook, inserts sample data, adds a column chart linked to that data, updates a data point, and saves the workbook as ChartReport.xlsx.
class ChartReportGenerator
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        // Add a column chart linked to the data range
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Update chart data (e.g., modify a value)
        sheet.Cells["B3"].PutValue(25); // change value for category B

        // Save the modified workbook to the specified file
        workbook.Save("ChartReport.xlsx", SaveFormat.Xlsx);
    }
}
