// Title: Define Print Area A1:G30 on the Active Worksheet with Aspose.Cells for .NET (C#)
// Description: This C# snippet creates a new Workbook, selects the first worksheet, assigns the PageSetup.PrintArea to the range A1:G30, and saves the file as PrintAreaDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# print area | Worksheet PageSetup | PrintArea property | A1:G30 range | save workbook | Aspose.Cells .NET API
// Common Searches: Aspose.Cells set print area C# | How to define printable range A1:G30 in Aspose.Cells | PageSetup.PrintArea example .NET | Set worksheet print area programmatically Aspose | C# Aspose.Cells print area before saving
// Developer Intent: Programmatically set the printable range of the active worksheet to cells A1 through G30.
// Use Cases: Generate a report where only the first 30 rows and 7 columns are printed. | Create a template that forces a fixed page layout for consistent printing across users. | Limit the printed output of a large workbook to a specific data block before distribution. | Prepare a spreadsheet for batch printing with a predefined area to reduce paper waste.
// AI Prompts: Write C# code that sets the print area to the used range of a worksheet with Aspose.Cells. | Show how to remove an existing PrintArea and assign a new range in Aspose.Cells for .NET. | Provide an example that configures margins, orientation, and a custom print area together using Aspose.Cells.

using System;
using Aspose.Cells;

namespace PrintAreaExample
{
    // This C# snippet creates a new Workbook, selects the first worksheet, assigns the PageSetup.PrintArea to the range A1:G30, and saves the file as PrintAreaDemo.xlsx using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first (active) worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the custom print area to cover cells A1:G30
            worksheet.PageSetup.PrintArea = "A1:G30";

            // Save the workbook (lifecycle: save)
            workbook.Save("PrintAreaDemo.xlsx");

            Console.WriteLine("Print area set to A1:G30 and workbook saved successfully.");
        }
    }
}
