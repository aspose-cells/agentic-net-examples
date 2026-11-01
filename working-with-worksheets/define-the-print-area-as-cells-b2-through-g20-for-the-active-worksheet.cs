// Title: Define a print area of B2:G20 on the active worksheet with Aspose.Cells for .NET
// AI Prompts: Assign "B2:G20" to the active worksheet's PageSetup.PrintArea property and then save the workbook as an .xlsx file. | Instantiate a Workbook, retrieve its active Worksheet, set the print range to B2:G20 via PageSetup, and write the file to disk.
// Common Searches: Aspose.Cells .NET how to set a custom print area for the active sheet | C# programmatically define worksheet print range B2 to G20 before saving | using PageSetup.PrintArea to specify print area in Aspose.Cells
// Tags: Aspose.Cells set print area .NET | C# PageSetup.PrintArea usage | worksheet print range B2:G20 example | save workbook after configuring print area

using Aspose.Cells;
using System;

// Creates a new workbook, accesses the active worksheet, sets its PageSetup.PrintArea to B2:G20, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the active worksheet
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Define the print area as B2:G20
        activeSheet.PageSetup.PrintArea = "B2:G20";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
