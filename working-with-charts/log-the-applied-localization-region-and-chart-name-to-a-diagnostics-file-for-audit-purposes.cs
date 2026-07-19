// Title: Log workbook region and chart title to a diagnostics file using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, sets its Region to Japan, adds a column chart named "Sales Chart", then appends the region and chart title to diagnostics.txt before saving the file as Output.xlsx.
// Keywords: Aspose.Cells | C# | .NET | log workbook region | chart title logging | diagnostics file | regional setting | audit trail | Excel chart metadata
// Common Searches: Aspose.Cells log workbook region to text file | C# write chart title to diagnostics file | How to record Excel chart metadata with Aspose.Cells | Save workbook locale and chart name for audit | Append region and chart title to log using Aspose.Cells
// Developer Intent: Append the workbook's regional setting and the chart's title to a diagnostics file for audit purposes.
// Use Cases: Maintain an audit log of locale information for workbooks generated in a multi‑region reporting pipeline. | Track chart identifiers alongside regional settings to simplify compliance verification. | Automate troubleshooting by recording workbook metadata each time a report is produced.
// AI Prompts: Generate C# code with Aspose.Cells that logs the workbook Region and every chart Title to a CSV file, including timestamps. | Show a thread‑safe way to append multiple chart entries to a diagnostics log using Aspose.Cells for .NET. | Provide a script that reads diagnostics.txt and produces a summary of region usage across generated workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, sets its Region to Japan, adds a column chart named "Sales Chart", then appends the region and chart title to diagnostics.txt before saving the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the regional setting for the workbook (example: Japan)
        workbook.Settings.Region = CountryCode.Japan;

        // Prepare sample data for the chart
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Set a title for the chart (used as the chart name for logging)
        chart.Title.Text = "Sales Chart";

        // Log the applied region and chart title to a diagnostics file
        string diagnosticsFile = "diagnostics.txt";
        string logEntry = $"Region: {workbook.Settings.Region}, ChartTitle: {chart.Title.Text}{Environment.NewLine}";
        File.AppendAllText(diagnosticsFile, logEntry);

        // Save the workbook
        workbook.Save("Output.xlsx");
    }
}
