// Title: Define Print Area B2:G20 on Active Worksheet with Aspose.Cells for .NET (C#)
// Description: Shows how to assign the print area to cells B2 through G20 on the active worksheet using Aspose.Cells for .NET, then save the workbook.
// Keywords: Aspose.Cells | .NET | C# | print area | PageSetup | B2:G20 | worksheet print range | save workbook
// Common Searches: Aspose.Cells set print area C# | How to define print range B2:G20 Aspose.Cells | C# PageSetup PrintArea Aspose.Cells example | Set active worksheet print area Aspose.Cells .NET | Save workbook after setting print area Aspose
// Developer Intent: Set the active worksheet's print area to cells B2 through G20.
// Use Cases: Print only a data table (B2:G20) while keeping other worksheet content off the page. | Create an invoice template that consistently prints the invoice section defined by B2:G20. | Automate batch processing of workbooks to restrict each file to a predefined printable range before distribution.
// AI Prompts: Generate C# code to set the same print area on all worksheets in an Aspose.Cells workbook. | Provide a script that reads the current print area, expands it to include newly added rows, and saves the file. | Explain how to combine page margins, orientation, and print area settings using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to assign the print area to cells B2 through G20 on the active worksheet using Aspose.Cells for .NET, then save the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the active worksheet (first worksheet by default)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the print area from cell B2 to G20
        worksheet.PageSetup.PrintArea = "B2:G20";

        // Save the workbook to a file
        workbook.Save("PrintAreaDemo.xlsx");

        Console.WriteLine("Print area set to B2:G20 and workbook saved.");
    }
}
