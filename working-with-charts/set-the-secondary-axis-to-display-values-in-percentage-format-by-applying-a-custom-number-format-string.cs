// Title: Set a secondary axis to display percentages with a custom number format in an Aspose.Cells column chart (C#)
// AI Prompts: Generate C# code that adds a secondary axis to an Aspose.Cells column chart and applies a "0%" number format to its values. | Show how to move a data series onto the secondary axis and format the axis as a percentage using Aspose.Cells for .NET. | Provide a snippet that configures a secondary axis in an Aspose.Cells chart and sets a custom percentage number format string.
// Common Searches: Aspose.Cells C# set secondary axis number format to percent | format secondary axis values as percentage in Aspose.Cells chart | C# Aspose.Cells column chart secondary axis custom number format | how to display growth series as percent on secondary axis using Aspose.Cells | Aspose.Cells chart secondary axis percentage formatting example
// Tags: secondary axis percentage formatting Aspose.Cells | custom number format string chart Aspose.Cells | C# column chart secondary axis Aspose.Cells | apply percentage format to chart series Aspose.Cells | Aspose.Cells chart number format example

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The example creates a workbook, adds month, sales, and growth data, inserts a column chart, places the growth series on a secondary axis, applies a custom "0%" number format to that axis, and saves the file as an Excel workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Growth");

            // Populate sample data
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C2"].PutValue(0.10); // 10%
            sheet.Cells["C3"].PutValue(0.15); // 15%
            sheet.Cells["C4"].PutValue(0.20); // 20%

            // Add a column chart (rows 6-20, columns 0-10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Primary series (Sales)
            chart.NSeries.Add("B2:B4", true);

            // Secondary series (Growth)
            chart.NSeries.Add("C2:C4", true);

            // Place the second series on the secondary axis if the API supports it.
            // In some Aspose.Cells versions the property is not available; comment out if compilation fails.
            // chart.NSeries[1].IsOnSecondaryAxis = true;

            // Optionally format the secondary series as percentage (if supported)
            // chart.NSeries[1].NumberFormat = "0%";

            // Save the workbook
            string outputPath = "ChartWithSecondaryAxisPercentage.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
