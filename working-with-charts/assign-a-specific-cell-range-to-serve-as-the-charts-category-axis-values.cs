// Title: Set Category Axis Range for a Column Chart with Aspose.Cells (C# .NET)
// Description: Creates a workbook, fills column A with labels and column B with values, adds a column chart, binds the series to B2:B10, and assigns the horizontal axis labels by setting NSeries.CategoryData to A2:A10, then saves the file as an XLSX document.
// Keywords: Aspose.Cells chart category axis | C# set chart categories | .NET Aspose.Cells example | NSeries.CategoryData property | column chart from worksheet cells | Excel chart category range | Aspose.Cells API usage
// Common Searches: Aspose.Cells set chart category axis range C# | How to bind chart categories to cells in Aspose.Cells | NSeries CategoryData example .NET | Assign cell range to chart axis Aspose.Cells | Create column chart with custom categories using Aspose.Cells
// Developer Intent: Assign a specific worksheet cell range as the category (horizontal) axis values for a chart created with Aspose.Cells.
// Use Cases: Display month names stored in column A on the X‑axis of a sales column chart. | Generate a product‑category performance chart that updates automatically when the source range changes. | Build a dashboard where the axis labels are driven by a named range, allowing non‑developers to modify categories directly in Excel.
// AI Prompts: Provide C# code that sets the CategoryData property of a chart series to a worksheet range using Aspose.Cells. | Show how to change both the series values and the category labels for a line chart in Aspose.Cells. | Explain how to refresh the CategoryData range after inserting new rows into the source data with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCategoryAxisDemo
{
    // Creates a workbook, fills column A with labels and column B with values, adds a column chart, binds the series to B2:B10, and assigns the horizontal axis labels by setting NSeries.CategoryData to A2:A10, then saves the file as an XLSX document.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: column A for categories, column B for values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 10; i++)
            {
                sheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("=Sheet1!$B$2:$B$10", true);

            // Assign a specific cell range as the category axis values
            // This uses the SeriesCollection.CategoryData property
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$10";

            // Save the workbook to an XLSX file
            workbook.Save("CategoryAxisDemo.xlsx");
        }
    }
}
