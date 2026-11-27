// Title: Dynamically set each worksheet's print area to its used range with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells set print area to used range for each worksheet | how to assign PageSetup.PrintArea based on MaxDisplayRange in Aspose.Cells | auto adjust Excel print area before saving workbook using Aspose.Cells .NET | programmatically define print area for all sheets in an Excel file with Aspose.Cells | Aspose.Cells C# calculate used range and set worksheet print area dynamically
// Tags: Aspose.Cells PageSetup.PrintArea API | C# MaxDisplayRange to PrintArea conversion | dynamic worksheet print area calculation | batch set Excel print area Aspose.Cells | auto adjust used range print area .NET

using System;
using System.IO;
using Aspose.Cells;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The example loads an existing workbook, iterates through every worksheet, obtains the used range via MaxDisplayRange, builds the address string (e.g., "A1:D10"), assigns it to the worksheet's PageSetup.PrintArea, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Iterate through each worksheet and set its print area dynamically
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range of the worksheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                // Calculate the start and end cell addresses of the used range
                string startCell = CellsHelper.CellIndexToName(usedRange.FirstRow, usedRange.FirstColumn);
                string endCell = CellsHelper.CellIndexToName(
                    usedRange.FirstRow + usedRange.RowCount - 1,
                    usedRange.FirstColumn + usedRange.ColumnCount - 1);

                // Build the print area string (e.g., "A1:D10")
                string printArea = $"{startCell}:{endCell}";

                // Assign the calculated print area to the worksheet
                sheet.PageSetup.PrintArea = printArea;
            }

            // Save the workbook after processing
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
