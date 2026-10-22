// Title: How to write a numeric value to cell B2 using A1 notation and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a new Workbook, accesses cell B2 via its A1 address, assigns a decimal number, and saves the file as Output.xlsx using Aspose.Cells. | Show an example of using Worksheet.Cells["B2"] to put a numeric value into the cell and then persisting the workbook in .NET with Aspose.Cells. | Provide a minimal Aspose.Cells snippet that demonstrates creating a workbook, writing 123.45 to B2, and calling Workbook.Save to produce an XLSX file.
// Common Searches: Aspose.Cells C# how to set a decimal value in cell B2 using A1 style | example of writing numeric data to a specific cell and saving the workbook with Aspose.Cells .NET | using Worksheet.Cells indexer to update a cell value and export to XLSX in C#
// Tags: set cell value using A1 address Aspose.Cells | write numeric data to worksheet cell C# Aspose.Cells | save workbook as XLSX Aspose.Cells .NET | create new workbook and modify cell Aspose.Cells | Worksheet.Cells indexer example Aspose.Cells

using Aspose.Cells;
using System;

// // Creates a workbook, accesses cell B2 via A1 notation, assigns a numeric value, and saves the file as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Access cell B2 using its A1 style name and set a numeric value
        Cell cell = sheet.Cells["B2"];
        cell.PutValue(123.45); // Example numeric value

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
