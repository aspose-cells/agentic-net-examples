// Title: Save a workbook with a customized column chart as an XLSX file using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a column chart, sets its title, and saves the workbook to an XLSX file with Aspose.Cells. | Write a .NET console example that adds sample data, builds a chart, and calls Workbook.Save with SaveFormat.Xlsx. | Show how to export a chart‑enhanced Excel workbook to .xlsx using the Aspose.Cells API.
// Common Searches: how to save an Aspose.Cells workbook that contains a chart to xlsx in C# | Aspose.Cells example for exporting a column chart to an Excel file | C# code to create a chart and write it to an .xlsx file with Aspose.Cells
// Tags: Aspose.Cells save workbook with chart to XLSX | C# create column chart Aspose.Cells | Workbook.Save using SaveFormat.Xlsx Aspose.Cells | set chart title Aspose.Cells C# | export chart‑enhanced Excel file .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveDemo
{
    // Demonstrates creating a new workbook, adding sample data, inserting a column chart with a title, and saving the file as CustomizedChart.xlsx (XLSX) using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Initialize a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["A4"].PutValue("Cherries");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Customize the chart (example: set title)
            chart.Title.Text = "Fruit Sales";

            // Save the workbook with the customized chart to XLSX format
            workbook.Save("CustomizedChart.xlsx", SaveFormat.Xlsx);

            // Optional: inform the user
            Console.WriteLine("Workbook with customized chart saved as CustomizedChart.xlsx");
        }
    }
}
