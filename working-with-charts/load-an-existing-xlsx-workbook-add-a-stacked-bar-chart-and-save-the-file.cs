// Title: Add a Stacked Bar Chart to an Existing XLSX Workbook with Aspose.Cells for .NET (C#)
// Description: Loads an existing XLSX file, accesses the first worksheet, inserts a BarStacked chart (rows 5‑20, columns 1‑8) linked to a specified data range, and saves the workbook as a new file containing the chart.
// Keywords: Aspose.Cells stacked bar chart C# | add BarStacked chart Aspose | load Excel workbook Aspose.Cells | save workbook with chart .NET | Excel chart automation C# | Aspose.Cells chart example
// Common Searches: how to insert a stacked bar chart into an existing Excel file using Aspose.Cells | Aspose.Cells C# example for adding BarStacked chart | create and save Excel chart programmatically with Aspose | Aspose.Cells chart types tutorial | stacked bar chart Aspose.Cells .NET guide
// Developer Intent: Programmatically add a stacked bar chart to a pre‑existing XLSX workbook and write the updated file.
// Use Cases: Enhance a monthly sales template by loading it, adding a stacked bar chart that compares product categories, and exporting the finished report. | Automate dashboard generation where each worksheet receives a BarStacked chart based on pre‑populated metrics before distribution. | Build a reusable utility method that accepts input and output paths, injects a stacked bar chart into the first sheet, and returns the modified workbook.
// AI Prompts: Generate a C# method that takes input and output XLSX paths, loads the workbook with Aspose.Cells, adds a BarStacked chart to the first worksheet using a given range, and saves the result. | Show sample code to add multiple stacked bar charts to different worksheets in the same workbook, each referencing its own data range, using Aspose.Cells for .NET. | Explain how to customize the appearance of a BarStacked chart (colors, legend, title, data labels) after creating it with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing XLSX file, accesses the first worksheet, inserts a BarStacked chart (rows 5‑20, columns 1‑8) linked to a specified data range, and saves the workbook as a new file containing the chart.
class AddStackedBarChart
{
    static void Main()
    {
        // Load the existing workbook from disk
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // OPTIONAL: If the workbook does not already contain data, you can add sample data here.
        // sheet.Cells["A1"].PutValue("Category");
        // sheet.Cells["B1"].PutValue("Series1");
        // sheet.Cells["C1"].PutValue("Series2");
        // for (int i = 2; i <= 5; i++)
        // {
        //     sheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
        //     sheet.Cells[$"B{i}"].PutValue(i * 10);
        //     sheet.Cells[$"C{i}"].PutValue(i * 15);
        // }

        // Add a stacked bar chart (BarStacked) to the worksheet.
        // Parameters: ChartType, topRow, leftColumn, bottomRow, rightColumn (zero‑based indices)
        int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 1, 20, 8);
        Chart chart = sheet.Charts[chartIndex];

        // Define the data range for the chart.
        // Adjust the range to match the actual data in your workbook.
        chart.NSeries.Add("=Sheet1!$A$1:$C$5", true);

        // Save the workbook with the new chart.
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
