// Title: Assign chart category labels from cells B2:B8 using Aspose.Cells NSeries.CategoryData in C#
// AI Prompts: Create a C# example that adds a column chart and assigns its CategoryData property to the range B2:B8 with Aspose.Cells. | Write code to bind chart series category labels to worksheet cells using the NSeries.CategoryData API in Aspose.Cells for .NET.
// Common Searches: how to set chart category axis labels from a worksheet range using Aspose.Cells in C# | Aspose.Cells NSeries.CategoryData property example for column charts | link Excel chart categories to cells B2:B8 with Aspose.Cells .NET library
// Tags: Aspose.Cells NSeries.CategoryData binding | C# column chart category labels from worksheet range | set chart categories using Aspose.Cells API | Excel chart series category data linking .NET | Aspose.Cells chart category axis configuration

using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, fills cells A2:A8 and B2:B8 with sample data, adds a column chart, sets its series values to B2:B8, links the chart's CategoryData to the same range, and saves the file as ChartCategoryData.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (optional, just to have something in the range)
        for (int i = 2; i <= 8; i++)
        {
            sheet.Cells[$"A{i}"].PutValue($"Category {i - 1}");
            sheet.Cells[$"B{i}"].PutValue(i * 10);
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the series values (vertical range)
        chart.NSeries.Add("B2:B8", true);

        // Link the category labels to the same range B2:B8
        chart.NSeries.CategoryData = "B2:B8";

        // Save the workbook
        workbook.Save("ChartCategoryData.xlsx");
    }
}
