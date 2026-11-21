// Title: Create an Excel workbook, fill cells with data, and save it as a tab‑delimited TXT file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export worksheet to tab delimited text file | How to generate a TSV file from a workbook with Aspose.Cells .NET | Save Excel data as a .txt file with tabs using Aspose.Cells in C# | Programmatically create a workbook and output tab‑separated values using Aspose.Cells | C# example for SaveFormat.TabDelimited to produce a text file
// Tags: Aspose.Cells generate tab-separated text | programmatic workbook creation C# | worksheet cell population Aspose.Cells | save workbook as TSV .NET | SaveFormat.TabDelimited usage

using System;
using Aspose.Cells;

// The sample creates a new Workbook, adds a header row and three data rows to the first worksheet, and then saves the workbook as a tab‑delimited TXT file (ExportedData.txt) using Aspose.Cells' SaveFormat.TabDelimited.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate data programmatically
        // Header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Country");

        // Sample data rows
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["C2"].PutValue("USA");

        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["C3"].PutValue("Canada");

        sheet.Cells["A4"].PutValue("Charlie");
        sheet.Cells["B4"].PutValue(28);
        sheet.Cells["C4"].PutValue("UK");

        // Export the workbook as a tab‑delimited TXT file
        workbook.Save("ExportedData.txt", SaveFormat.TabDelimited);
    }
}
