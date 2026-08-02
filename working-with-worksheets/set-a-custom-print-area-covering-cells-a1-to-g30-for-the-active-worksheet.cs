// Title: Define a custom print area (A1:G30) for the active worksheet with Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, selects the first worksheet, sets PageSetup.PrintArea to "A1:G30", and saves the file, showing how to define a custom print range for the active sheet using Aspose.Cells.
// Keywords: Aspose.Cells print area | C# set print range | PageSetup.PrintArea | custom print area A1:G30 | Aspose.Cells worksheet printing | define print area .NET
// Common Searches: Aspose.Cells set print area C# | How to define print range A1:G30 Aspose.Cells | C# print only selected cells Aspose.Cells | PageSetup.PrintArea example | Set worksheet print area programmatically
// Developer Intent: The developer wants to define a custom print area covering cells A1 through G30 on the active worksheet.
// Use Cases: Generate printable reports that include only a specific data block. | Avoid blank pages by limiting the printed region to the needed range. | Create reusable templates where only a defined range is printed, ignoring other worksheet content.
// AI Prompts: Provide C# code to set the print area dynamically based on the used range in Aspose.Cells. | Show how to clear an existing print area and assign a new range on a different worksheet using Aspose.Cells. | Explain how to apply the same print area to all worksheets in a workbook with Aspose.Cells.

using System;
using Aspose.Cells;

namespace PrintAreaDemo
{
    // This C# example creates a workbook, selects the first worksheet, sets PageSetup.PrintArea to "A1:G30", and saves the file, showing how to define a custom print range for the active sheet using Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first (active) worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the custom print area to cover cells A1 through G30
            worksheet.PageSetup.PrintArea = "A1:G30";

            // Save the workbook to a file
            workbook.Save("PrintAreaDemo.xlsx");
        }
    }
}
