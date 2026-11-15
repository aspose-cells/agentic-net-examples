// Title: How to set rows 1‑3 as repeating print titles on every printed page with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# repeat first three rows on each printed page | set print title rows 1 to 3 using Aspose.Cells .NET | configure worksheet page setup to repeat header rows in Excel with C# | C# Aspose.Cells print titles for rows 1-3 across all pages | Excel repeat rows on every page programmatically using Aspose.Cells
// Tags: set print title rows Aspose.Cells | repeat header rows Excel C# | worksheet page setup repeat rows | Aspose.Cells print titles Excel | C# configure print titles worksheet

using Aspose.Cells;

// The example creates or loads a workbook, accesses the first worksheet, assigns "$1:$3" to the PageSetup.PrintTitleRows property so rows 1‑3 repeat on every printed page, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // create new workbook

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set rows 1 to 3 to repeat on every printed page
        sheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
