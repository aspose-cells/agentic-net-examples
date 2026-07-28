// Title: Create a Workbook and add a single Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to instantiate an Aspose.Cells Workbook, add one worksheet, optionally rename it, and save the file as an XLSX document using C#.
// Keywords: Aspose.Cells | C# workbook creation | add worksheet | worksheet rename | save as xlsx | Excel generation .NET | Aspose.Cells example
// Common Searches: Aspose.Cells add worksheet C# | how to create a workbook with Aspose.Cells .NET | set worksheet name Aspose.Cells example | save workbook as xlsx using Aspose.Cells
// Developer Intent: Generate an Excel file, insert a single sheet, give it a custom name, and write it to disk.
// Use Cases: Initialize a blank report workbook before populating data. | Create a template file that contains a named sheet for later chart insertion. | Produce an empty Excel package for downstream automation pipelines.
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a worksheet called "Data", and saves it as "Report.xlsx". | Show how to add multiple worksheets with distinct names using Aspose.Cells for .NET. | Explain how to change a worksheet's visibility and tab color after adding it with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to instantiate an Aspose.Cells Workbook, add one worksheet, optionally rename it, and save the file as an XLSX document using C#.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Add a single worksheet to the workbook
        int newSheetIndex = workbook.Worksheets.Add();

        // Optionally set a name for the added worksheet
        workbook.Worksheets[newSheetIndex].Name = "AddedSheet";

        // Save the workbook to verify the result
        workbook.Save("AddedWorksheet.xlsx", SaveFormat.Xlsx);
    }
}
