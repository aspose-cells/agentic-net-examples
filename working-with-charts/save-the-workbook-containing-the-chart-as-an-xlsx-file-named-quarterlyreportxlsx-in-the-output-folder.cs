// Title: Save a Workbook with a Column Chart to QuarterlyReport.xlsx using Aspose.Cells for .NET
// Description: Creates a new workbook, adds quarter‑wise sales data, inserts a column chart, ensures an "output" folder exists, and saves the file as QuarterlyReport.xlsx (XLSX) with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | save workbook | column chart | XLSX | output folder | create directory | Excel export | .NET chart example | QuarterlyReport.xlsx
// Common Searches: Aspose.Cells save workbook with chart C# | How to export chart to XLSX using Aspose.Cells .NET | Create output directory and save Excel file Aspose.Cells | Save Excel file to specific path C# Aspose.Cells | Generate quarterly sales chart and save as XLSX
// Developer Intent: Generate a workbook that contains a column chart and write it to the file QuarterlyReport.xlsx in an "output" directory.
// Use Cases: Automated quarterly sales reporting with a visual column chart. | Batch generation of Excel files with embedded charts for distribution. | Creating Excel templates in memory and persisting them to a known folder for downstream processing. | Saving chart‑enabled workbooks to a server location for web applications.
// AI Prompts: Provide C# code that adds a line chart to a workbook and saves it as SalesTrend.xlsx in a "reports" folder using Aspose.Cells. | Explain how to set a chart title, axis labels, and legend before saving the workbook with Aspose.Cells. | Show how to export the same workbook to PDF and PNG formats in addition to XLSX using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds quarter‑wise sales data, inserts a column chart, ensures an "output" folder exists, and saves the file as QuarterlyReport.xlsx (XLSX) with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Quarter");
        worksheet.Cells["A2"].PutValue("Q1");
        worksheet.Cells["A3"].PutValue("Q2");
        worksheet.Cells["A4"].PutValue("Q3");
        worksheet.Cells["A5"].PutValue("Q4");

        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["B2"].PutValue(120);
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["B4"].PutValue(130);
        worksheet.Cells["B5"].PutValue(170);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data source for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Ensure the output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Define the full path for the XLSX file
        string filePath = Path.Combine(outputDir, "QuarterlyReport.xlsx");

        // Save the workbook as XLSX using the provided Save method
        workbook.Save(filePath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook with chart saved to: {filePath}");
    }
}
