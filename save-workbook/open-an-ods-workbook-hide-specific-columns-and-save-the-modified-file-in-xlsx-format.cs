// Title: Load an ODS workbook, hide specific columns, and save as XLSX using Aspose.Cells for .NET (C#)
// Common Searches: asp.net hide column B and D in an ODS file before converting to XLSX with Aspose.Cells | c# hide multiple columns in ODS workbook using Aspose.Cells and export to Excel format | how to programmatically hide columns in an ODS spreadsheet and save as XLSX in .NET | Aspose.Cells hide column by index in ODS worksheet then save as Xlsx | convert ODS to XLSX while removing column visibility using C# Aspose.Cells
// Tags: hide columns Aspose.Cells C# | ODS to XLSX conversion Aspose.Cells | worksheet column visibility Aspose.Cells | save workbook as Xlsx Aspose.Cells | column index hide Aspose.Cells | modify ODS worksheet Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an ODS workbook, hides columns B and D on the first worksheet, and saves the modified workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the ODS workbook
        Workbook workbook = new Workbook("input.ods");

        // Get the first worksheet (you can change the index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Hide specific columns.
        // Example: hide column B (index 1) and column D (index 3)
        sheet.Cells.HideColumn(1); // Hide column B
        sheet.Cells.HideColumn(3); // Hide column D

        // Save the modified workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
