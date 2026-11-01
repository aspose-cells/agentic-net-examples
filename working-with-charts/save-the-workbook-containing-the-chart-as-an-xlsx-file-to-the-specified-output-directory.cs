// Title: Save an Aspose.Cells workbook containing a column chart to a chosen directory as an XLSX file using C#
// AI Prompts: Generate C# code that creates a new workbook, fills sample data, adds a column chart, and saves the file as XLSX to a user‑specified folder with Aspose.Cells. | Adapt the example to insert a line chart, name the output file with the current date, and write it to a configurable path using Aspose.Cells in .NET.
// Common Searches: asp.net aspocells create column chart and save workbook to custom folder | c# how to export an Excel file with a chart to a specific directory using Aspose.Cells | save workbook with chart as xlsx programmatically Aspose.Cells .NET | specify output path when saving Aspose.Cells workbook containing a chart | example of adding a chart to a worksheet and writing the file to C:\Data folder with Aspose.Cells
// Tags: create column chart Aspose.Cells C# | save workbook as xlsx Aspose.Cells | specify output directory Aspose.Cells | add chart to worksheet Aspose.Cells | export chart workbook .NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The program builds a new workbook, populates cells A1:B4 with sample data, inserts a column chart based on the values, and saves the workbook as 'ChartWorkbook.xlsx' in the specified output directory.
class Program
{
    static void Main()
    {
        try
        {
            // Output directory for the generated workbook
            string outputDir = @"C:\Output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true); // Values series
            // Category data can be set via CategoryData property if available.
            // If the property is not present in the used version, the chart will use default categories.

            // Build full file path
            string filePath = Path.Combine(outputDir, "ChartWorkbook.xlsx");

            // Save the workbook
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to: {filePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
