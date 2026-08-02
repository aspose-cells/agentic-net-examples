// Title: Aspose.Cells for .NET – Set Print Area B2:G20 on Active Worksheet (C#)
// Description: Learn how to define the printable region B2:G20 on the active worksheet with Aspose.Cells for .NET. The example creates a workbook, sets Worksheet.PageSetup.PrintArea, and saves the file, showing how to persist a custom print area in C#.
// Keywords: Aspose.Cells print area | C# set worksheet print range | PageSetup.PrintArea | B2:G20 Excel region | .NET Aspose.Cells printable area | define print area programmatically | Excel print settings C#
// Common Searches: Aspose.Cells how to set print area in C# | Set worksheet print range B2 to G20 using Aspose.Cells | C# Aspose.Cells PageSetup PrintArea example | Programmatically limit Excel print area with Aspose.Cells
// Developer Intent: Define the printable region B2:G20 for the active worksheet using Aspose.Cells.
// Use Cases: Generate reports where only a specific data block (B2:G20) should be printed. | Create templates that automatically restrict printing to a predefined area. | Prepare workbooks for batch printing with consistent printable zones across sheets.
// AI Prompts: Show code to set a custom print area for several worksheets with Aspose.Cells in C#. | How can I read the current PrintArea of a worksheet and adjust it based on user input? | Combine margin, orientation, and print area settings in Aspose.Cells PageSetup.

using System;
using Aspose.Cells;

// Learn how to define the printable region B2:G20 on the active worksheet with Aspose.Cells for .NET. The example creates a workbook, sets Worksheet.PageSetup.PrintArea, and saves the file, showing how to persist a custom print area in C#.
class SetPrintAreaDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first (active) worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the print area to cells B2 through G20
        worksheet.PageSetup.PrintArea = "B2:G20";

        // Save the workbook (optional, demonstrates that the setting is persisted)
        workbook.Save("PrintAreaDemo.xlsx");
    }
}
