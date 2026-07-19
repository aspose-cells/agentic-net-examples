// Title: Aspose.Cells for .NET: Link Chart Category Labels to Cells Using CategoryData (C#)
// Description: Demonstrates creating a workbook, populating categories and values, adding a column chart, setting the series range, and binding the chart's CategoryData property to the cell range B2:B8 so the category axis reflects worksheet data. The file is saved as CategoryDataLinked.xlsx.
// Keywords: Aspose.Cells | C# chart CategoryData | set chart categories Aspose.Cells | link chart axis to cells | column chart Aspose.Cells .NET | CategoryData property example | Excel chart automation
// Common Searches: Aspose.Cells set CategoryData range C# | How to bind chart categories to worksheet cells Aspose.Cells | Assign category axis values in Aspose.Cells chart | C# Aspose.Cells chart category labels from range | Link Excel chart categories programmatically
// Developer Intent: Use the CategoryData property to bind a chart’s category axis to the cell range B2:B8 in an Aspose.Cells .NET workbook.
// Use Cases: Generate Excel reports with dynamic chart categories that update when source cells change | Create quick prototypes where the same range supplies both series values and category labels | Automate dashboards that require column charts driven by worksheet data
// AI Prompts: Show how to set CategoryData to A2:A8 while keeping series values in B2:B8 in Aspose.Cells C# | Provide code for multiple series each with different value ranges but sharing a single CategoryData range | Explain how to refresh chart categories after modifying source cells at runtime using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, populating categories and values, adding a column chart, setting the series range, and binding the chart's CategoryData property to the cell range B2:B8 so the category axis reflects worksheet data. The file is saved as CategoryDataLinked.xlsx.
class SetCategoryDataDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (optional, just to have something in the range)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 8; i++)
        {
            sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series data range (vertical series)
        chart.NSeries.Add("B2:B8", true);

        // Define the category labels by linking CategoryData to the range B2:B8
        chart.NSeries.CategoryData = "B2:B8";

        // Save the workbook to a file
        workbook.Save("CategoryDataLinked.xlsx");
    }
}
