// Title: Dynamically set each worksheet's print area to its used range with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set worksheet print area to used range programmatically | How to assign print area based on MaxDisplayRange for all sheets in Aspose.Cells | C# loop through worksheets and set PageSetup.PrintArea before saving | Set dynamic print area for an Excel workbook using Aspose.Cells .NET | Update print area to actual data range in Aspose.Cells before printing
// Tags: Aspose.Cells worksheet print area automation | C# set PageSetup.PrintArea using MaxDisplayRange | dynamic Excel print area with Aspose.Cells | loop over worksheets assign print area .NET | used range based print area Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an Excel workbook, iterates each worksheet, obtains its MaxDisplayRange, builds an A1‑style address, assigns it to ws.PageSetup.PrintArea, and saves the file, ensuring the print area matches the used range for every sheet.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Set print area for each worksheet based on its used range
            foreach (Worksheet ws in workbook.Worksheets)
            {
                try
                {
                    // Get the used range of the worksheet
                    AsposeRange usedRange = ws.Cells.MaxDisplayRange;

                    // Ensure the range contains at least one cell
                    if (usedRange != null && usedRange.RowCount > 0 && usedRange.ColumnCount > 0)
                    {
                        // Build the address in A1 notation (e.g., "A1:C10")
                        string startCell = CellsHelper.CellIndexToName(usedRange.FirstRow, usedRange.FirstColumn);
                        string endCell = CellsHelper.CellIndexToName(
                            usedRange.FirstRow + usedRange.RowCount - 1,
                            usedRange.FirstColumn + usedRange.ColumnCount - 1);
                        string printArea = $"{startCell}:{endCell}";

                        // Assign the print area
                        ws.PageSetup.PrintArea = printArea;
                    }
                }
                catch (Exception ex)
                {
                    // Handle errors for individual worksheets without stopping the whole process
                    Console.WriteLine($"Warning: Could not set print area for worksheet \"{ws.Name}\": {ex.Message}");
                }
            }

            // Save the workbook with the updated print areas
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
