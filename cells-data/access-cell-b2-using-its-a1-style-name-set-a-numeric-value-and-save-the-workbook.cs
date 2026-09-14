// Title: How to set a numeric value in cell B2 using A1 notation and save the workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a new Workbook, accesses cell B2 via its A1 address, assigns the number 123.45, and saves the file as Output.xlsx using Aspose.Cells. | Demonstrate using Aspose.Cells' PutValue method to write a decimal value to a specific cell identified by A1 notation and then persist the workbook.
// Common Searches: Aspose.Cells C# set value in B2 cell using A1 address | How to write a decimal to a specific cell with Aspose.Cells .NET | Saving an Excel workbook after updating a cell with Aspose.Cells | C# Aspose.Cells PutValue example for numeric data | Create new workbook and write number to B2 using Aspose.Cells
// Tags: putvalue numeric cell Aspose.Cells C# | write value to cell B2 Aspose.Cells | save workbook as xlsx Aspose.Cells | access cell by A1 address Aspose.Cells | create new workbook Aspose.Cells .NET

using System;
using Aspose.Cells;

// // Creates a new workbook, writes the numeric value 123.45 to cell B2 using A1 notation, and saves the file as Output.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell B2 using its A1 style name and set a numeric value
        worksheet.Cells["B2"].PutValue(123.45);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
