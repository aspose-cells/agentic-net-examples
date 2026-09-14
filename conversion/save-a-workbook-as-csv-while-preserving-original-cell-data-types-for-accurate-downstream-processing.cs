// Title: Save an Aspose.Cells workbook as a CSV file while preserving numeric and date cell types in C#
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, adds mixed text, numeric, and date values, and saves the active worksheet to a CSV file using TxtSaveOptions with a comma delimiter. | Show how to configure TxtSaveOptions to export only the active sheet, keep the workbook in memory after saving, and retain original data types by setting PreserveString to false. | Demonstrate disabling ClearData, setting the CSV separator, and saving the file without losing the original numeric and date formats.
// Common Searches: Aspose.Cells C# export active worksheet to CSV with original number formatting | How to keep date values when saving Excel as CSV using Aspose.Cells .NET | TxtSaveOptions CSV separator comma preserve data types Aspose.Cells example | Save workbook as CSV without clearing data from memory Aspose.Cells C#
// Tags: Aspose.Cells CSV export with TxtSaveOptions | preserve numeric and date types in CSV export | export active worksheet to CSV .NET | configure CSV separator comma Aspose.Cells | prevent workbook data clearance after CSV save

using System;
using Aspose.Cells;
using System.Text;

// Creates a workbook, adds text, numeric, and date values, configures TxtSaveOptions (CSV format, comma separator, ExportAllSheets = false, ClearData = false, PreserveString = false) and saves the active sheet to output.csv, preserving the original cell data types for downstream processing.
class SaveWorkbookAsCsv
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Ensure values are stored with their original data types
        cells.PreserveString = false; // default, but set explicitly for clarity

        // Add header row
        cells["A1"].PutValue("Name");
        cells["B1"].PutValue("Score");
        cells["C1"].PutValue("Date");

        // Add sample data with mixed types
        cells["A2"].PutValue("Alice");
        cells["B2"].PutValue(95.5);                     // numeric
        cells["C2"].PutValue(new DateTime(2023, 5, 1)); // date

        cells["A3"].PutValue("Bob");
        cells["B3"].PutValue(88);                       // numeric
        cells["C3"].PutValue(new DateTime(2023, 5, 2)); // date

        // Configure CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        csvOptions.Separator = ',';          // comma delimiter
        csvOptions.ClearData = false;        // keep workbook in memory after saving
        csvOptions.ExportAllSheets = false;  // export only the active sheet

        // Save the workbook as CSV while preserving original cell data types
        workbook.Save("output.csv", csvOptions);
    }
}
