// Title: How to set rows 1‑3 as repeating print titles on every page using Aspose.Cells for .NET (C#)
// AI Prompts: Apply the PageSetup.PrintTitleRows property to make rows 1‑3 repeat on each printed page in a C# workbook with Aspose.Cells. | Define a print title range of $1:$3 for a worksheet and save the workbook using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells C# set print title rows 1 to 3 for each printed page | repeat header rows on every printed page using Aspose.Cells PageSetup | define print title range $1:$3 in Aspose.Cells .NET | C# Aspose.Cells page setup repeat rows across printed pages | set worksheet print titles for rows 1‑3 with Aspose.Cells API
// Tags: Aspose.Cells page setup repeat rows | worksheet print titles configuration .NET | Excel A1 style range $1:$3 Aspose | C# set repeating header rows for printing | configure print title rows Aspose.Cells

using Aspose.Cells;

// The example creates or loads a workbook, accesses a worksheet, sets the PageSetup.PrintTitleRows property to "$1:$3" so that rows 1 through 3 repeat on every printed page, and saves the workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet (or any specific worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Configure rows 1 to 3 to repeat on every printed page
        // The range is specified in Excel A1 style with absolute references
        sheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
