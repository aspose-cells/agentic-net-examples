// Title: Aspose.Cells .NET: Create 10 Worksheets with Individual Column Charts in a Loop
// Description: This example shows how to build a Workbook, clear the default sheet, and use a for‑loop to add ten worksheets (Sheet1‑Sheet10). Each sheet receives sample category/value data, a column chart positioned between rows 7‑25 and columns A‑I, and the chart is bound to the sheet's own data range (B2:B5) and categories (A2:A5). The workbook is saved as TenSheetsWithCharts.xlsx.
// Keywords: Aspose.Cells | C# | .NET | multiple worksheets | column chart | chart binding | Excel automation | loop | Workbook.Save | NSeries | CategoryData | generate Excel file | TenSheetsWithCharts
// Common Searches: Add a chart to every worksheet using Aspose.Cells C# | Create multiple Excel sheets with individual charts in .NET | Bind chart series to sheet-specific range Aspose.Cells | Loop to generate worksheets with charts Aspose.Cells | Save workbook with ten charts C#
// Developer Intent: Programmatically generate a workbook containing ten worksheets, each with its own column chart linked to that sheet's data.
// Use Cases: Regional sales dashboard where each region gets its own sheet and column chart. | Automated financial reporting workbook with separate sheets for quarters, each showing a chart. | Test data generation for QA that requires several sheets with distinct charts. | Presentation deck export where each slide is a worksheet with a chart. | Educational sample demonstrating chart binding in Aspose.Cells.
// AI Prompts: Generate C# code with Aspose.Cells that creates 10 worksheets, adds sample data, and inserts a column chart bound to that sheet's data. | Explain step‑by‑step how to bind NSeries and CategoryData to ranges on the current worksheet in Aspose.Cells. | Refactor the loop into a reusable method that returns a worksheet with a chart, then call it ten times. | Provide a GitHub‑style README snippet describing the TenSheetsWithCharts example. | Show how to customize chart titles and styles for each sheet in a loop using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example shows how to build a Workbook, clear the default sheet, and use a for‑loop to add ten worksheets (Sheet1‑Sheet10). Each sheet receives sample category/value data, a column chart positioned between rows 7‑25 and columns A‑I, and the chart is bound to the sheet's own data range (B2:B5) and categories (A2:A5). The workbook is saved as TenSheetsWithCharts.xlsx.
    public class TenSheetsWithCharts
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (default contains one worksheet)
                Workbook workbook = new Workbook();

                // Remove the default worksheet to start clean
                workbook.Worksheets.Clear();

                // Loop to create 10 worksheets, each with its own data and chart
                for (int i = 1; i <= 10; i++)
                {
                    // Add a new worksheet with a distinct name
                    Worksheet sheet = workbook.Worksheets.Add($"Sheet{i}");

                    // Populate sample data: header + 4 rows of categories and values
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");

                    // Example categories: A, B, C, D
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["A4"].PutValue("C");
                    sheet.Cells["A5"].PutValue("D");

                    // Values vary per sheet to illustrate uniqueness
                    sheet.Cells["B2"].PutValue(i * 10 + 1);
                    sheet.Cells["B3"].PutValue(i * 10 + 2);
                    sheet.Cells["B4"].PutValue(i * 10 + 3);
                    sheet.Cells["B5"].PutValue(i * 10 + 4);

                    // Add a column chart to the worksheet
                    // Position the chart from row 7, column 0 to row 25, column 8
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
                    Chart chart = sheet.Charts[chartIndex];

                    // Bind the chart to the data range of the current sheet
                    string dataRange = $"{sheet.Name}!B2:B5";
                    string categoryRange = $"{sheet.Name}!A2:A5";

                    chart.NSeries.Add(dataRange, true);
                    chart.NSeries.CategoryData = categoryRange;

                    // Optional: set a title for clarity
                    chart.Title.Text = $"Chart for {sheet.Name}";
                }

                // Define output file path
                string outputPath = "TenSheetsWithCharts.xlsx";

                // Save the workbook with all worksheets and charts
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TenSheetsWithCharts.Run();
        }
    }
}
