// Title: Insert a Timestamp Worksheet into a Merged Excel Workbook Using Aspose.Cells for .NET
// Common Searches: how to insert a timestamp sheet into an existing merged Excel file with Aspose.Cells C# | Aspose.Cells add new worksheet with current date and time after merging workbooks | C# write merge operation datetime to a new sheet in an Excel workbook using Aspose.Cells | save merged workbook with audit timestamp worksheet using Aspose.Cells for .NET | Aspose.Cells create timestamp worksheet programmatically in C#
// Tags: insert audit timestamp sheet Aspose.Cells | populate cell with current datetime C# | create additional worksheet programmatically Aspose.Cells | track merge operation in Excel workbook | export merged workbook with extra sheet Aspose.Cells

using Aspose.Cells;
using System;

// The example loads a merged Excel file, adds a new worksheet named "Timestamp", writes the current date and time (formatted as yyyy-MM-dd HH:mm:ss) into cell A1, optionally adds a label in B1, and saves the workbook as a new file that includes the timestamp sheet.
class Program
{
    static void Main()
    {
        // Load the merged workbook
        Workbook workbook = new Workbook("MergedWorkbook.xlsx");

        // Add a new worksheet for the timestamp
        int sheetIndex = workbook.Worksheets.Add();
        Worksheet timestampSheet = workbook.Worksheets[sheetIndex];
        timestampSheet.Name = "Timestamp";

        // Write the current date and time into cell A1
        timestampSheet.Cells["A1"].PutValue(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

        // Optional: add a descriptive label in B1
        timestampSheet.Cells["B1"].PutValue("Merge operation timestamp");

        // Save the workbook with the new timestamp sheet
        workbook.Save("MergedWorkbook_WithTimestamp.xlsx");
    }
}
