// Title: Aspose.Cells .NET: Create a Column Chart and Save the Workbook as XLSX
// Description: C# example that builds a new Workbook, fills cells A1:B4 with categories and values, adds a column chart linked to that range, and persists the file using Workbook.Save with SaveFormat.Xlsx.
// Keywords: Aspose.Cells .NET | create column chart | add chart to worksheet | save workbook as xlsx | Workbook.Save | SaveFormat.Xlsx | C# Excel chart example | Aspose.Cells chart API | export chart to Excel | chart generation C#
// Common Searches: Aspose.Cells add column chart C# | save workbook with chart Aspose.Cells | C# Aspose.Cells chart to XLSX | how to export chart using Aspose.Cells .NET | Aspose.Cells chart example save as xlsx
// Developer Intent: Generate an Excel file that contains a programmatically created column chart.
// Use Cases: Produce a sales‑by‑region column chart and deliver the workbook to clients in XLSX format. | Automate nightly generation of product performance charts and archive the resulting XLSX files. | Build a web service that returns a customized Excel dashboard with charts created on the fly.
// AI Prompts: Show how to change the column chart to a line chart and then save the workbook as XLSX with Aspose.Cells. | Provide code to add a chart title, axis labels, and a legend before persisting the workbook. | Explain how to resize and reposition the chart programmatically and then export the workbook to an XLSX file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// C# example that builds a new Workbook, fills cells A1:B4 with categories and values, adds a column chart linked to that range, and persists the file using Workbook.Save with SaveFormat.Xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (uses the Workbook() constructor rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["A4"].PutValue("Cherry");

        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(30);
        worksheet.Cells["B3"].PutValue(45);
        worksheet.Cells["B4"].PutValue(25);

        // Add a column chart to the worksheet
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);          // Values
        chart.NSeries.CategoryData = "A2:A4";      // Categories

        // Save the workbook containing the chart as an XLSX file
        // (uses the Save(string, SaveFormat) rule)
        workbook.Save("ChartWorkbook.xlsx", SaveFormat.Xlsx);
    }
}
