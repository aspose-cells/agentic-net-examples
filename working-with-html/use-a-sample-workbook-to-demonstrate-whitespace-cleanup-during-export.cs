// Title: C# Aspose.Cells CSV Export: Trim or Preserve Leading Blank Rows/Columns
// Description: This example creates a workbook, writes data starting at cell C3 (leaving empty rows and columns at the top‑left), and saves two CSV files. One file keeps the blanks (TrimLeadingBlankRowAndColumn = false) and the other removes them (TrimLeadingBlankRowAndColumn = true) using TxtSaveOptions.
// Keywords: Aspose.Cells CSV export | TrimLeadingBlankRowAndColumn | remove leading blanks Aspose.Cells | preserve whitespace CSV | TxtSaveOptions C#
// Common Searches: Aspose.Cells how to trim leading blank rows when saving to CSV | C# export Excel to CSV without empty rows | TxtSaveOptions TrimLeadingBlankRowAndColumn example | remove empty columns Aspose.Cells CSV | preserve whitespace in CSV export Aspose.Cells
// Developer Intent: Demonstrate controlling whitespace trimming during CSV export with Aspose.Cells.
// Use Cases: Generate a CSV that mirrors the original worksheet layout, including leading empty rows/columns, for legacy systems that rely on fixed positions. | Create a compact CSV by stripping unnecessary leading blanks to lower file size and speed up downstream processing. | Produce side‑by‑side CSV files (trimmed vs. untrimmed) to evaluate the effect of whitespace on data‑parsing scripts.
// AI Prompts: Write C# code with Aspose.Cells to export a worksheet to CSV while keeping all leading empty rows and columns. | Extend the sample to also trim trailing blank rows and columns during CSV export. | Explain the impact of the TrimLeadingBlankRowAndColumn property on the output of TxtSaveOptions.

using System;
using Aspose.Cells;

// This example creates a workbook, writes data starting at cell C3 (leaving empty rows and columns at the top‑left), and saves two CSV files. One file keeps the blanks (TrimLeadingBlankRowAndColumn = false) and the other removes them (TrimLeadingBlankRowAndColumn = true) using TxtSaveOptions.
class Program
{
    static void Main()
    {
        try
        {
            WhitespaceCleanupDemo.Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

class WhitespaceCleanupDemo
{
    public static void Run()
    {
        // Initialize a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate cells starting from C3, leaving leading blank rows and columns
        worksheet.Cells["C3"].PutValue("Data1");
        worksheet.Cells["D4"].PutValue("Data2");
        worksheet.Cells["E5"].PutValue("Data3");

        // Export without trimming leading blank rows/columns
        TxtSaveOptions saveOptions = new TxtSaveOptions();
        saveOptions.TrimLeadingBlankRowAndColumn = false; // keep blanks
        workbook.Save("output_with_blanks.csv", saveOptions);

        // Export with trimming leading blank rows/columns
        saveOptions.TrimLeadingBlankRowAndColumn = true; // remove blanks
        workbook.Save("output_trimmed.csv", saveOptions);
    }
}
