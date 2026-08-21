// Title: Aspose.Cells for .NET – Set numeric value in B2 (A1 notation) and save workbook
// Description: Create a new Workbook, access cell B2 using A1 notation, assign the number 123.45, and save the file as Output.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# set cell value | A1 notation B2 Aspose.Cells | save Excel workbook Aspose.Cells | write numeric data to Excel cell | Aspose.Cells example .NET
// Common Searches: Aspose.Cells write number to B2 C# | How to save Excel file after editing with Aspose.Cells | Access cells by A1 style in Aspose.Cells .NET | C# Aspose.Cells set cell value and save
// Developer Intent: Create a workbook, put a numeric value into cell B2 using A1 notation, and persist the workbook to disk.
// Use Cases: Generate a financial report where the total amount is placed in B2 before exporting. | Automate data entry by programmatically inserting calculated figures into a known cell location. | Initialize a template workbook, populate a key metric in B2, and deliver the file to downstream systems.
// AI Prompts: Write C# code with Aspose.Cells that sets 123.45 in cell B2 (A1 notation) and saves as Output.xlsx. | Show how to format the value in B2 as currency using Aspose.Cells before saving the workbook. | Demonstrate writing several numeric values to different cells using A1 notation and then saving the workbook.

using System;
using Aspose.Cells;

// Create a new Workbook, access cell B2 using A1 notation, assign the number 123.45, and save the file as Output.xlsx with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell B2 using A1 notation and set a numeric value
        worksheet.Cells["B2"].PutValue(123.45);

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
