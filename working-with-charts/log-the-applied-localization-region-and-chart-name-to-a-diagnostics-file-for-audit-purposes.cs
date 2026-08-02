// Title: Log workbook region and chart title to an audit file with Aspose.Cells for .NET
// Description: This C# example creates a workbook, sets its Region to Japan, adds a column chart named "Quarterly Sales", and writes a timestamp, the workbook's regional setting, and the chart title to an audit.log file before saving the workbook.
// Keywords: Aspose.Cells | .NET | log workbook region | chart title audit | diagnostics file | CountryCode Japan | Excel chart logging | audit log C#
// Common Searches: Aspose.Cells log workbook region to file | How to write chart title to audit log in C# | Save Excel workbook regional settings with Aspose.Cells | Create diagnostics log for Aspose.Cells charts | C# audit trail for generated Excel files
// Developer Intent: Record the workbook's regional setting and the chart's title in a log file for compliance, monitoring, or troubleshooting purposes.
// Use Cases: Generate a compliance‑ready audit trail each time a workbook is produced, capturing region and chart identifiers. | Facilitate troubleshooting by logging timestamps, regional settings, and chart names for every automated Excel export. | Maintain a central diagnostics file for multiple charts by iterating through a workbook and appending each chart's title and region.
// AI Prompts: Write C# code that loops through all charts in an Aspose.Cells workbook and logs each chart's title with the workbook's Region to a CSV audit file. | Provide a reusable method that records timestamp, region, and chart name to a configurable log path, including error handling for file I/O. | Show how to convert the CountryCode enum value to its localized display name and include it in the audit log entry.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# example creates a workbook, sets its Region to Japan, adds a column chart named "Quarterly Sales", and writes a timestamp, the workbook's regional setting, and the chart title to an audit.log file before saving the workbook.
class AuditDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Set the regional settings for the workbook
        workbook.Settings.Region = CountryCode.Japan; // Example region

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(150);
        sheet.Cells["B3"].PutValue(250);

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set chart title (this will be used as the chart name for auditing)
        chart.Title.Text = "Quarterly Sales";

        // Bind data to the chart
        chart.NSeries.Add("B2:B3", true);
        chart.NSeries.CategoryData = "A2:A3";

        // Log the applied region and chart name to a diagnostics file
        string logFilePath = "audit.log";
        using (StreamWriter writer = new StreamWriter(logFilePath, true))
        {
            writer.WriteLine($"Timestamp: {DateTime.Now:O}");
            writer.WriteLine($"Workbook Region: {workbook.Settings.Region}");
            writer.WriteLine($"Chart Name: {chart.Title.Text}");
            writer.WriteLine(new string('-', 40));
        }

        // Save the workbook to a file
        workbook.Save("AuditDemo.xlsx");
    }
}
