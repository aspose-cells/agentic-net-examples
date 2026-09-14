// Title: How to set rows 1‑2 as repeating print titles on a worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Configure Aspose.Cells PageSetup to repeat rows 1 through 2 as print titles on every printed page in a C# workbook. | Generate an Excel file where the first two rows are defined as repeating header rows using the PrintTitleRows property of Aspose.Cells. | Apply worksheet print title settings in C# so that rows 1‑2 appear on each printed page with Aspose.Cells.
// Common Searches: Aspose.Cells C# repeat first two rows on each printed page | set print title rows $1:$2 using Aspose.Cells PageSetup | how to make header rows repeat when printing Excel with Aspose.Cells .NET | C# Aspose.Cells example for PrintTitleRows property | PageSetup.PrintTitleRows usage in Aspose.Cells workbook
// Tags: Aspose.Cells PageSetup PrintTitleRows | repeat header rows Excel C# | worksheet print titles Aspose.Cells | C# generate Excel with repeating rows | set print title rows Aspose.Cells workbook

using Aspose.Cells;

// The sample creates a new Workbook, accesses the first Worksheet, sets PageSetup.PrintTitleRows to "$1:$2" so rows 1‑2 repeat on each printed page, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Set rows 1 through 2 to repeat on each printed page
        sheet.PageSetup.PrintTitleRows = "$1:$2";

        // Save the workbook
        workbook.Save("output.xlsx");
    }
}
