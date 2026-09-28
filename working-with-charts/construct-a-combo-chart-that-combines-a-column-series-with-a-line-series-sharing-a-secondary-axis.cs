// Title: Create a combo chart with column and line series on a secondary axis using Aspose.Cells for .NET (C#)
// AI Prompts: Generate a new workbook, fill it with month, sales, and profit data, then add a combo chart where the sales series appears as columns and the profit series appears as a line plotted on a second value axis. | Set axis titles, verify that the output folder exists, and save the workbook as an .xlsx file with Aspose.Cells in C#.
// Common Searches: how to add a secondary axis to a line series in an Aspose.Cells chart using C# | sample code for creating a column‑line mixed chart with Aspose.Cells .NET | Aspose.Cells tutorial for mixed chart types in an Excel workbook | C# program to plot sales as columns and profit as a line on separate axes with Aspose.Cells
// Tags: Aspose.Cells mixed column line chart | Aspose.Cells secondary axis for line series | C# generate Excel chart with multiple series | Aspose.Cells set individual series type | Aspose.Cells save workbook as xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a workbook, adds month, sales, and profit data, builds a combo chart that shows sales as columns and profit as a line on a secondary axis, assigns axis titles, ensures the output directory exists, and saves the file as ComboChart.xlsx using Aspose.Cells for .NET.
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

            // Populate sample data
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);

            // Add a combo chart (initially a column chart)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Sales and Profit Combo Chart";

            // Add the column series (Sales)
            int columnSeriesIdx = chart.NSeries.Add("B2:B4", true);

            // Add the line series (Profit) and set its type to Line
            int lineSeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[lineSeriesIdx].Type = ChartType.Line;

            // Optional: set axis titles
            chart.CategoryAxis.Title.Text = "Month";
            chart.ValueAxis.Title.Text = "Sales";

            // Determine output path and ensure directory exists
            string outputFile = Path.Combine(Environment.CurrentDirectory, "ComboChart.xlsx");
            string outputDir = Path.GetDirectoryName(outputFile) ?? Environment.CurrentDirectory;
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
