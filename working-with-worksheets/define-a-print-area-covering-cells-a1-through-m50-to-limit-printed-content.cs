// Title: Set a custom print area (A1:M50) on the first worksheet using Aspose.Cells in C#
// AI Prompts: Generate C# code that assigns the PageSetup.PrintArea property to "A1:M50" for a worksheet and saves the workbook. | Show how to limit the printable region of an Excel sheet to cells A1 through M50 with Aspose.Cells. | Provide a step‑by‑step example that creates a new workbook, defines a print range, and writes the file to disk.
// Common Searches: Aspose.Cells C# how to define a print area for a worksheet | set printable range A1:M50 using Aspose.Cells .NET API | programmatically limit Excel print output to specific cells in C# | configure PageSetup.PrintArea in Aspose.Cells example
// Tags: Aspose.Cells printable region setup | C# PageSetup.PrintArea property | Excel worksheet printable range definition | save workbook after configuring print area

using System;
using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, sets the printable area to cells A1‑M50 via PageSetup.PrintArea, and saves the file as PrintAreaExample.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty Excel file)
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Define the print area to cover cells A1 through M50
        sheet.PageSetup.PrintArea = "A1:M50";

        // Save the workbook to a file
        workbook.Save("PrintAreaExample.xlsx");
    }
}
