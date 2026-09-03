// Title: Set columns A and B as repeating print titles in an Aspose.Cells worksheet using C#
// AI Prompts: Generate C# code that configures Aspose.Cells PageSetup to repeat columns A and B as print titles on every printed page. | Show how to assign the PrintTitleColumns property to "$A:$B" and then save the workbook with Aspose.Cells. | Provide a complete example that creates a workbook, sets print title columns, and verifies the setting in C#.
// Common Searches: asp.net aspose.cells set print title columns a:b | repeat first two columns on each printed page using Aspose.Cells C# | how to configure worksheet print titles programmatically in .NET | C# Aspose.Cells PageSetup PrintTitleColumns example | save Excel file after setting print title columns with Aspose.Cells
// Tags: Aspose.Cells worksheet PrintTitleColumns property | C# repeat column headers on printed pages | Aspose.Cells PageSetup configuration for Excel export | set print titles for Excel workbook using .NET | save workbook after configuring print titles Aspose.Cells

using Aspose.Cells;

// // Creates a new workbook, sets columns A and B as print title columns so they repeat on each printed page, and saves the file as PrintTitleColumns.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set print title columns to columns A and B (repeated on each printed page)
        sheet.PageSetup.PrintTitleColumns = "$A:$B";

        // Save the workbook
        workbook.Save("PrintTitleColumns.xlsx");
    }
}
