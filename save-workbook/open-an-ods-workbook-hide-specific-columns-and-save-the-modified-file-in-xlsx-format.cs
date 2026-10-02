// Title: Hide specific columns in an ODS workbook and export to XLSX using Aspose.Cells for .NET
// AI Prompts: Programmatically hide column B and column D in an ODS file and then save the workbook as an XLSX file with Aspose.Cells for .NET. | Convert an ODS spreadsheet to XLSX while making selected columns invisible (e.g., B and D) using C# and the Aspose.Cells API. | Load an ODS workbook, hide chosen columns on the first worksheet, and output the result as an XLSX document via Aspose.Cells.
// Common Searches: Aspose.Cells C# hide columns B and D in ODS before converting to XLSX | How to hide specific columns in an ODS workbook using Aspose.Cells for .NET | Convert ODS to XLSX with hidden columns using Aspose.Cells example in C# | C# code to hide columns in first worksheet of ODS and save as XLSX
// Tags: hide columns in ODS workbook Aspose.Cells | ODS to XLSX conversion with column visibility control | Aspose.Cells column hiding API C# | modify worksheet column visibility ODS Aspose | export ODS as XLSX with hidden columns Aspose.Cells

using System;
using Aspose.Cells;

// Loads an ODS workbook, hides columns B and D on the first worksheet, and saves the modified file as XLSX using Aspose.Cells for .NET.
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
