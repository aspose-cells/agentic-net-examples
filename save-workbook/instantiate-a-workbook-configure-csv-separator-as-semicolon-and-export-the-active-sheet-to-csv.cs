// Title: Save only the active worksheet as a semicolon‑delimited CSV file with Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, sets TxtSaveOptions.Separator to ';', and saves only the active sheet to a CSV file using Aspose.Cells. | Show how to configure Aspose.Cells TxtSaveOptions to export a single worksheet as a semicolon‑separated CSV in a .NET application.
// Common Searches: how to use Aspose.Cells to save only the active sheet as a CSV with a semicolon delimiter | C# Aspose.Cells export a specific worksheet to CSV using a custom separator | setting TxtSaveOptions separator property for CSV output in Aspose.Cells | export current worksheet to semicolon‑separated CSV file with Aspose.Cells .NET
// Tags: Aspose.Cells TxtSaveOptions custom CSV delimiter | save active sheet as CSV Aspose.Cells | single worksheet CSV export .NET | semicolon delimited CSV with Aspose.Cells | configure CSV separator in Aspose.Cells

using System;
using Aspose.Cells;

// The program creates a new workbook, fills the first worksheet with sample data, marks it as the active sheet, configures TxtSaveOptions to use a semicolon as the CSV separator and to export only the active sheet, then saves the result as ActiveSheet.csv.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle create rule)
        Workbook workbook = new Workbook();

        // Populate the active worksheet with sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Alice");
        sheet.Cells["B3"].PutValue(25);

        // Ensure the first worksheet is the active one
        workbook.Worksheets.ActiveSheetIndex = 0;

        // Configure CSV save options: use semicolon as separator
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ';';            // Set delimiter to semicolon
        csvOptions.ExportAllSheets = false;    // Export only the active sheet

        // Save the active sheet to a CSV file (lifecycle save rule)
        workbook.Save("ActiveSheet.csv", csvOptions);
    }
}
