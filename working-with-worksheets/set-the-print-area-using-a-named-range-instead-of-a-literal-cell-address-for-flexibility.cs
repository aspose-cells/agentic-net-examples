// Title: Set a worksheet's print area using a named range with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add sample data, define a named range (MyPrintArea) covering A1:B3, assign that name to PageSetup.PrintArea, and save the file, enabling flexible print‑area configuration.
// Keywords: Aspose.Cells | C# | print area | named range | PageSetup.PrintArea | worksheet printing | dynamic print range | .NET Excel library
// Common Searches: Aspose.Cells set print area named range C# | how to use named range for print area in Aspose.Cells | PageSetup.PrintArea from workbook name Aspose.Cells | C# Aspose.Cells flexible print area example
// Developer Intent: Set the worksheet's print area by referencing a named range instead of a hard‑coded cell address.
// Use Cases: Update the printable region by changing the named range without touching code. | Apply a consistent print layout across multiple sheets by reusing the same named range. | Adjust the print area dynamically based on data volume before exporting or printing.
// AI Prompts: Write C# code with Aspose.Cells that creates a named range and uses it as the worksheet's print area. | Show how to modify an existing named range and refresh the PrintArea property in an Aspose.Cells workbook. | Provide a script that sets the same named‑range print area for every worksheet in a workbook using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsPrintAreaWithNamedRange
{
    // Demonstrates how to create a workbook, add sample data, define a named range (MyPrintArea) covering A1:B3, assign that name to PageSetup.PrintArea, and save the file, enabling flexible print‑area configuration.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(20);

            // Define a named range that covers A1:B3
            // Add a new name to the workbook's name collection
            int nameIndex = workbook.Worksheets.Names.Add("MyPrintArea");
            // The RefersTo formula must start with '=' and include the sheet name
            workbook.Worksheets.Names[nameIndex].RefersTo = $"={sheet.Name}!$A$1:$B$3";

            // Set the print area to the named range (instead of a literal address)
            sheet.PageSetup.PrintArea = "MyPrintArea";

            // Save the workbook (lifecycle: save)
            workbook.Save("PrintAreaWithNamedRange.xlsx");

            Console.WriteLine("Workbook saved with print area set to named range 'MyPrintArea'.");
        }
    }
}
