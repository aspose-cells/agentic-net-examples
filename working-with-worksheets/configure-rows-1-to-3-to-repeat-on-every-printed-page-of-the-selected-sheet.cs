// Title: Set rows 1‑3 to repeat as print titles on each printed page of the selected worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, uses Aspose.Cells to make rows 1 through 3 repeat as print titles on every printed page of the active worksheet, and saves the workbook. | Show how to configure the PageSetup.PrintTitleRows property in Aspose.Cells so specific rows are repeated when printing an Excel sheet. | Demonstrate applying a repeat‑rows page‑setup setting to the currently selected worksheet and exporting the result with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# set rows 1 to 3 as print titles for active worksheet | repeat header rows on each printed page using Aspose.Cells .NET | how to use PageSetup.PrintTitleRows in Aspose.Cells for Excel printing | C# Aspose.Cells configure rows to repeat on every printed page
// Tags: Aspose.Cells PageSetup.PrintTitleRows | repeat rows on printed pages .NET | set print title rows Excel Aspose.Cells | active worksheet page setup Aspose.Cells | C# Excel repeat header rows Aspose

using Aspose.Cells;

// Loads an Excel workbook, accesses the active worksheet, assigns rows 1‑3 to the PrintTitleRows property so they repeat on every printed page, and saves the modified file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your file path)
        var workbook = new Workbook("input.xlsx");

        // Get the currently selected worksheet
        var sheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // Configure rows 1 to 3 to repeat on every printed page
        sheet.PageSetup.PrintTitleRows = "$1:$3";

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
