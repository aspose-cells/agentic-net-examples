// Title: Apply an AutoFilter to a header row and freeze the first row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example for adding an AutoFilter to the first row and freezing it | how to keep filter dropdowns visible by freezing header row in Aspose.Cells | C# code to set AutoFilter range and freeze panes in an Excel workbook with Aspose.Cells | freeze top row while using AutoFilter in Aspose.Cells for .NET
// Tags: autofilter header row Aspose.Cells | freeze panes first row C# | set autofilter range Aspose.Cells .NET | keep filter controls visible Aspose.Cells | worksheet freeze top row Aspose.Cells

using Aspose.Cells;
using System;

// Creates a new workbook, adds a header row with sample data, applies an AutoFilter to the range A1:C4, freezes the first row so filter controls stay visible, and saves the file as AutoFilterAndFreeze.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["C1"].PutValue("Country");

        // Add sample data
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["C2"].PutValue("USA");

        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(25);
        sheet.Cells["C3"].PutValue("UK");

        sheet.Cells["A4"].PutValue("Charlie");
        sheet.Cells["B4"].PutValue(28);
        sheet.Cells["C4"].PutValue("Canada");

        // Apply AutoFilter to the header row (first row)
        // The range includes all columns with data
        sheet.AutoFilter.Range = "A1:C4";

        // Freeze the header row so filter controls stay visible
        // Freeze first row (row index 1) and no columns
        sheet.FreezePanes(1, 0, 1, 0);

        // Save the workbook
        workbook.Save("AutoFilterAndFreeze.xlsx");
    }
}
