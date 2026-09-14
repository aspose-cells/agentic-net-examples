// Title: Hide row and column headings in printed Excel worksheets using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an existing .xlsx workbook, sets the worksheet's PageSetup.PrintHeadings to false, and saves the modified file. | Show how to configure the PageSetup of a specific worksheet in Aspose.Cells to prevent row and column headings from appearing in the print output. | Provide a step‑by‑step example that removes Excel print headings to create a cleaner layout for a marketing brochure using Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# hide row headings in printed Excel file | disable column and row headings in Excel print preview using Aspose.Cells | how to turn off PrintHeadings property for a worksheet in Aspose.Cells .NET | remove Excel print headings for brochure layout C# Aspose.Cells | Aspose.Cells page setup hide headings when exporting to PDF
// Tags: Aspose.Cells worksheet PageSetup PrintHeadings | C# hide Excel print headings | disable print headings Aspose.Cells | Excel brochure layout printing settings | Aspose.Cells .xlsx print configuration

using Aspose.Cells;

// // Loads an Excel workbook, disables printing of row and column headings on the first worksheet via PageSetup.PrintHeadings, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or specify the desired index/name)
        var worksheet = workbook.Worksheets[0];

        // Disable printing of row and column headings for a cleaner layout
        worksheet.PageSetup.PrintHeadings = false;

        // Save the workbook with the updated setting (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
