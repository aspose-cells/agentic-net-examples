// Title: How to set the print area to B2:G20 on the active worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Assign the PageSetup.PrintArea property to "B2:G20" on the first worksheet and save the workbook. | Configure a custom print range for the active sheet before exporting to Excel with Aspose.Cells in C#. | Programmatically define a print area for a worksheet and generate Result.xlsx using Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet print area to specific range B2:G20 | example of defining print area for active sheet using Aspose.Cells .NET | how to use PageSetup.PrintArea property in Aspose.Cells C# | set print range before saving workbook with Aspose.Cells | C# Aspose.Cells set print area for first worksheet
// Tags: Aspose.Cells set worksheet print area | PageSetup.PrintArea configuration | custom print area B2:G20 | Aspose.Cells worksheet print range | export with defined print area

using Aspose.Cells;

// Creates a new workbook, accesses the first worksheet, sets its print area to B2:G20 via the PageSetup.PrintArea property, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the print area as cells B2 through G20
        worksheet.PageSetup.PrintArea = "B2:G20";

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
