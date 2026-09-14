// Title: Open an ODS workbook, hide specific columns, and save it as XLSX with Aspose.Cells for .NET
// AI Prompts: Load an ODS file, hide columns B and C on the first worksheet, then export the workbook to XLSX using Aspose.Cells in C#. | Use Aspose.Cells to programmatically hide a range of columns in an ODS spreadsheet before converting it to XLSX. | Hide selected columns in an ODS workbook and save the modified file as XLSX with the Aspose.Cells API for .NET.
// Common Searches: Aspose.Cells hide columns B and C in ODS before converting to XLSX C# | C# how to hide columns in an ODS workbook using Aspose.Cells | convert ODS to XLSX while removing column visibility with Aspose.Cells | worksheet.Cells.HideColumns example for ODS files in .NET
// Tags: hide columns ODS Aspose.Cells | ODS to XLSX conversion C# | worksheet.Cells.HideColumns ODS | modify column visibility Aspose.Cells .NET | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an ODS workbook, hides columns B and C on the first worksheet, and saves the result as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the ODS workbook from file
        Workbook workbook = new Workbook("input.ods");

        // Get the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide specific columns.
        // Example: hide columns B and C (zero‑based indexes 1 and 2)
        // HideColumns(startColumnIndex, numberOfColumns)
        worksheet.Cells.HideColumns(1, 2);

        // Save the modified workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
