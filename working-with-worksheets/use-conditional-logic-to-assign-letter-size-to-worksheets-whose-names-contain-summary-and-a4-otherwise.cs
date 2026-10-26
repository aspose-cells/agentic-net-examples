// Title: C# – Set worksheet page setup paper size to Letter for sheets with “Summary” in the name and to A4 for all other worksheets using Aspose.Cells
// Common Searches: Aspose.Cells C# set paper size to Letter for worksheets whose name contains Summary | How to change Excel worksheet page setup to A4 unless the sheet name includes Summary using Aspose.Cells | Conditional page setup paper size based on worksheet name in a .NET workbook with Aspose.Cells | C# code to assign Letter paper size to summary tabs and A4 to remaining sheets in an Excel file | Aspose.Cells example for setting different paper sizes per worksheet programmatically
// Tags: Aspose.Cells conditional page setup | set worksheet paper size programmatically | letter paper size for specific sheets | a4 paper size default worksheet | C# workbook worksheet iteration

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, loops through each worksheet, and uses Aspose.Cells to set the PageSetup.PaperSize to PaperLetter when the worksheet name contains "Summary" (case‑insensitive); otherwise it sets the size to PaperA4, then saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet and set paper size based on worksheet name
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // If the worksheet name contains "Summary" (case‑insensitive), set Letter size; otherwise set A4
                if (sheet.Name.IndexOf("Summary", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperLetter;
                }
                else
                {
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
