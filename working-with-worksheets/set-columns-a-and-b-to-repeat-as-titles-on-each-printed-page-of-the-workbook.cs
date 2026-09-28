// Title: Repeat columns A and B as print titles on each printed page with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to assign "$A:$B" to the worksheet's PageSetup.PrintTitleColumns property. | Programmatically configure an Excel workbook so that columns A and B appear as titles on every printed page using Aspose.Cells.
// Common Searches: Aspose.Cells C# set print title columns A B | repeat first two columns on every printed page using Aspose.Cells | C# Aspose.Cells PageSetup.PrintTitleColumns example | configure worksheet to repeat columns when printing with Aspose.Cells | how to set print titles for columns in an Excel workbook using Aspose.Cells .NET
// Tags: Aspose.Cells PageSetup.PrintTitleColumns | repeat columns as print titles C# | Excel worksheet print title columns Aspose | C# set print titles for columns | Aspose.Cells workbook page setup repeat columns

using Aspose.Cells;

// Loads an Excel workbook, sets columns A and B to repeat as print titles on every printed page via the PageSetup.PrintTitleColumns property, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired worksheet index/name)
        var worksheet = workbook.Worksheets[0];

        // Set columns A and B to repeat as titles on each printed page
        worksheet.PageSetup.PrintTitleColumns = "$A:$B";

        // Save the workbook with the changes
        workbook.Save("output.xlsx");
    }
}
