// Title: Set the first two rows as repeatable print titles on each page using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to repeat rows 1 and 2 on every printed Excel page | Set PrintTitleRows property to $1:$2 with Aspose.Cells in .NET | Configure worksheet page setup to repeat header rows using Aspose.Cells C# example | Aspose.Cells repeat top rows when printing workbook programmatically | Using A1 notation to define print title rows in Aspose.Cells for .NET
// Tags: Aspose.Cells Worksheet PageSetup PrintTitleRows | C# repeat header rows on printed Excel pages | Excel A1 notation print title rows Aspose.Cells | Set print titles for workbook pages .NET | Aspose.Cells save workbook with repeated rows | Configure worksheet print settings programmatically

using Aspose.Cells;

// The example creates a Workbook, accesses the first Worksheet, sets the PageSetup.PrintTitleRows property to "$1:$2" so rows 1‑2 repeat on every printed page, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set rows 1 through 2 to repeat on each printed page
        // The range is specified in Excel A1 notation
        sheet.PageSetup.PrintTitleRows = "$1:$2";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
