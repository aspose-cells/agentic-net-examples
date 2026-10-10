// Title: Add a column chart and a pie chart with separate data ranges to the same worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to insert a Column chart referencing range B2:B5 and a Pie chart referencing range E2:E4 on a single worksheet, positioning them side by side. | Show how to assign titles and category data for multiple chart types in one Excel sheet with Aspose.Cells, then save the workbook.
// Common Searches: asp.net aspose.cells create multiple charts on one sheet | c# add column and pie charts to same worksheet with Aspose.Cells | how to set different data ranges for each chart in Aspose.Cells | position two charts next to each other in Excel using Aspose.Cells C# | Aspose.Cells chart NSeries category data example
// Tags: Aspose.Cells add multiple charts to worksheet | C# column chart data range Aspose.Cells | C# pie chart data range Aspose.Cells | Aspose.Cells chart positioning Excel | Aspose.Cells set chart titles programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart and ChartType

// The example creates a new workbook, populates two data tables, adds a column chart and a pie chart on the same worksheet with distinct data ranges and positions, sets chart titles, and saves the file as MultipleCharts.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the first chart (Column chart)
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

            // Populate data for the second chart (Pie chart)
            sheet.Cells["D1"].PutValue("Item");
            sheet.Cells["E1"].PutValue("Amount");
            sheet.Cells["D2"].PutValue("Apple");
            sheet.Cells["D3"].PutValue("Banana");
            sheet.Cells["D4"].PutValue("Cherry");
            sheet.Cells["E2"].PutValue(50);
            sheet.Cells["E3"].PutValue(30);
            sheet.Cells["E4"].PutValue(20);

            // Add the first chart - Column chart
            int chartIndex1 = sheet.Charts.Add(ChartType.Column, 7, 0, 22, 7);
            Chart chart1 = sheet.Charts[chartIndex1];
            chart1.NSeries.Add("=Sheet1!$B$2:$B$5", true);               // Values
            chart1.NSeries.CategoryData = "=Sheet1!$A$2:$A$5";          // Categories
            chart1.Title.Text = "Monthly Sales";

            // Add the second chart - Pie chart
            int chartIndex2 = sheet.Charts.Add(ChartType.Pie, 7, 9, 22, 16);
            Chart chart2 = sheet.Charts[chartIndex2];
            chart2.NSeries.Add("=Sheet1!$E$2:$E$4", true);               // Values
            chart2.NSeries.CategoryData = "=Sheet1!$D$2:$D$4";          // Categories
            chart2.Title.Text = "Fruit Distribution";

            // Save the workbook
            string outputPath = "MultipleCharts.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
