// Title: How to assign a worksheet cell range as the category axis for a column chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells set category axis range for column chart in C# | C# example using NSeries.CategoryData to define chart categories | Assign Excel range to chart categories with Aspose.Cells .NET | How to bind worksheet cells to chart category axis in Aspose.Cells | Create column chart with custom category labels using Aspose.Cells C#
// Tags: Aspose.Cells NSeries.CategoryData property | C# column chart category axis range | Aspose.Cells set chart categories from worksheet | programmatic Excel chart axis data range .NET | Aspose.Cells chart series values range example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, fills A2:A5 with category labels and B2:B5 with numeric values, adds a column chart, links the series to the values range, assigns the category axis to the range "Data!A2:A5" via the NSeries.CategoryData property, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and name it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for categories and values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series using the values range
            chart.NSeries.Add("Data!B2:B5", true);

            // Assign the category axis values (use NSeries.CategoryData)
            chart.NSeries.CategoryData = "Data!A2:A5";

            // Define output file path
            string outputPath = "ChartWithCategoryAxis.xlsx";

            // Save the workbook with the chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
