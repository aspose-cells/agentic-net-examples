// Title: Log workbook regional setting and chart name to a diagnostics file with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, sets its regional setting to Japan, adds a column chart named "SalesChart", saves the file, and appends a UTC timestamped entry containing the workbook's Region property and the chart's Name to a diagnostics.txt file for audit and compliance tracking.
// Keywords: Aspose.Cells | C# | .NET | workbook region logging | chart name audit | diagnostics file | regional settings Japan | Chart.Name | Settings.Region | File.AppendAllText | timestamped log
// Common Searches: Aspose.Cells log workbook region to file | record chart name in diagnostics with Aspose.Cells C# | audit Excel chart creation .NET | write localization info to text file using Aspose.Cells | timestamped chart metadata logging C#
// Developer Intent: Append the workbook's regional setting and the created chart's identifier to a diagnostics file for traceability.
// Use Cases: Compliance reporting: capture the locale (e.g., Japan) and chart identifier each time a workbook is generated. | Automated pipelines: generate a timestamped audit entry after saving a workbook to monitor regional compliance. | Multi‑workbook monitoring: aggregate diagnostics entries to analyze chart creation patterns across different locales.
// AI Prompts: Generate a reusable C# method that logs Aspose.Cells workbook Region and Chart.Name to a CSV file with error handling. | Show how to extend the audit log to include chart type, data range, and worksheet name using Aspose.Cells for .NET. | Provide a PowerShell script that reads the diagnostics.txt entries and summarizes chart creation counts per region.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example creates a workbook, sets its regional setting to Japan, adds a column chart named "SalesChart", saves the file, and appends a UTC timestamped entry containing the workbook's Region property and the chart's Name to a diagnostics.txt file for audit and compliance tracking.
class AuditChartCreation
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Set the regional settings for the workbook (e.g., Japan)
        wb.Settings.Region = CountryCode.Japan;

        // Get the first worksheet
        Worksheet sheet = wb.Worksheets[0];

        // Add some sample data for the chart
        sheet.Cells["A1"].PutValue("Month");
        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(120);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(180);

        // Create a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set chart data range and title
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";
        chart.Title.Text = "Quarterly Sales";

        // Assign a name to the chart for identification
        chart.Name = "SalesChart";

        // Save the workbook (lifecycle rule)
        wb.Save("AuditChart.xlsx");

        // Log the applied region and chart name to a diagnostics file
        string diagnosticsPath = "diagnostics.txt";
        string logEntry = $"Timestamp: {DateTime.UtcNow:u}, Region: {wb.Settings.Region}, ChartName: {chart.Name}";
        File.AppendAllText(diagnosticsPath, logEntry + Environment.NewLine);
    }
}
