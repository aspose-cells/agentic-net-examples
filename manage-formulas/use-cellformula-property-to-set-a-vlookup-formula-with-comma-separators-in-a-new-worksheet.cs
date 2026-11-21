// Title: Add a VLOOKUP formula with comma separators to a cell using Cell.Formula in Aspose.Cells for .NET
// Common Searches: how to assign a VLOOKUP formula to a cell with commas in Aspose.Cells C# | using Cell.Formula to insert VLOOKUP in a new worksheet with Aspose.Cells .NET | Aspose.Cells C# create lookup table and apply VLOOKUP formula programmatically | set VLOOKUP formula with comma delimiter in Excel file using Aspose.Cells API | save workbook after adding VLOOKUP formula with Aspose.Cells for .NET
// Tags: Cell.Formula set VLOOKUP Aspose.Cells | comma-separated formula syntax C# | create lookup table worksheet Aspose.Cells | save workbook as XLSX Aspose.Cells | programmatic VLOOKUP Excel C#

using Aspose.Cells;
using System;

// The example creates a new workbook, adds a worksheet named "Data", builds a lookup table in A1:B5, places the lookup key in C2, assigns a VLOOKUP formula with comma separators to D2 using the Cell.Formula property, and saves the file as VlookupExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add a new worksheet and name it
        Worksheet ws = workbook.Worksheets[workbook.Worksheets.Add()];
        ws.Name = "Data";

        // Populate a lookup table in A1:B5
        ws.Cells["A1"].PutValue("Key");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("Banana");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["A4"].PutValue("Cherry");
        ws.Cells["B4"].PutValue(30);
        ws.Cells["A5"].PutValue("Date");
        ws.Cells["B5"].PutValue(40);

        // Value to look up
        ws.Cells["C2"].PutValue("Banana");

        // Set VLOOKUP formula with comma separators using Cell.Formula
        ws.Cells["D2"].Formula = "VLOOKUP(C2, A1:B5, 2, FALSE)";

        // Save the workbook
        workbook.Save("VlookupExample.xlsx");
    }
}
