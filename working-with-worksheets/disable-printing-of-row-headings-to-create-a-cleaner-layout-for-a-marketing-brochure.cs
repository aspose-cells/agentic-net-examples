// Title: Turn off row and column headings when printing an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, disables printing of row and column headings, and saves it as an XLSX file. | Show how to configure PageSetup.PrintHeadings to false in Aspose.Cells to produce a clean printed layout. | Provide a minimal example that loads or creates a worksheet, turns off print headings, and saves the result for a marketing brochure.
// Common Searches: asp.net aspose.cells hide row headings in printed Excel sheet | c# disable column headings in page setup using Aspose.Cells | remove print headings from worksheet Aspose.Cells example | how to suppress Excel print headings with Aspose.Cells C#
// Tags: Aspose.Cells PageSetup.PrintHeadings false | disable print headings Excel .NET | clean worksheet layout for brochure Aspose.Cells | save workbook as XLSX without headings | printing settings worksheet Aspose.Cells

using Aspose.Cells;

// Creates or loads a workbook, accesses the first worksheet, sets PageSetup.PrintHeadings to false to omit row and column headings in printed output, and saves the file as MarketingBrochure.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Disable printing of row and column headings for a cleaner layout
        sheet.PageSetup.PrintHeadings = false;

        // Save the modified workbook
        workbook.Save("MarketingBrochure.xlsx", SaveFormat.Xlsx);
    }
}
