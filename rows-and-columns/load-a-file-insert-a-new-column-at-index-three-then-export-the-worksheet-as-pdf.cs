// Title: Insert a new column at position 4 in an Excel worksheet and save it as PDF using Aspose.Cells for .NET (C#)
// Common Searches: how to add a column before column D in an Excel file using Aspose.Cells C# | save a specific worksheet as PDF after inserting a column with Aspose.Cells | Aspose.Cells C# insert column at index 3 and export workbook to PDF | C# code to insert a column in Excel and convert the sheet to PDF using Aspose.Cells | Aspose.Cells example for inserting a column and generating PDF output
// Tags: Aspose.Cells insert column C# | Aspose.Cells save as PDF | Excel column insertion zero‑based index | worksheet PDF export Aspose.Cells | modify worksheet before PDF conversion

using Aspose.Cells;
using System;

// Loads input.xlsx, inserts a new column at zero‑based index 3 (before column D) in the first worksheet, and saves the workbook as output.pdf in PDF format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        var workbook = new Workbook("input.xlsx");

        // Get the first worksheet (you can change the index or name as needed)
        var worksheet = workbook.Worksheets[0];

        // Insert a new column at index 3 (zero‑based, so before column D)
        worksheet.Cells.InsertColumn(3);

        // Export the workbook (or the specific worksheet) as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
