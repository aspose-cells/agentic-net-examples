// Title: How to set rows 1‑3 as repeatable print titles on every page using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set first three rows as print titles for each printed page | repeat header rows on each page with Aspose.Cells PageSetup in .NET | C# code to assign PrintTitleRows property in Aspose.Cells workbook | configure page setup to repeat rows in Aspose.Cells for .NET | Aspose.Cells example for printing rows 1 to 3 on every page
// Tags: Aspose.Cells PageSetup PrintTitleRows property | C# configure print titles for rows | Excel workbook page setup with Aspose.Cells | Aspose.Cells .NET repeat rows on each printed page | set print title rows in Aspose.Cells workbook

using Aspose.Cells;

// The program creates a new workbook, accesses the first worksheet, assigns rows 1‑3 as print titles using the PageSetup.PrintTitleRows property, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the active worksheet (first worksheet by default)
        Worksheet sheet = workbook.Worksheets[0];

        // Set rows 1 to 3 as print titles so they repeat on each printed page
        sheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the workbook to a file
        workbook.Save("Output.xlsx");
    }
}
