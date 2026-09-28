// Title: Set worksheet print area dynamically to the used range with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that obtains the worksheet's MaxDisplayRange using Aspose.Cells and assigns the resulting A1 range to the PageSetup.PrintArea property. | Show how to translate the first and last cell indices of the used range into A1 notation and set the worksheet's print area before saving.
// Common Searches: Aspose.Cells C# set printable region to the worksheet's used range automatically | How to programmatically define the Excel print region based on the current data range in .NET | C# example for setting the worksheet's printable range to the data extent using Aspose.Cells
// Tags: Aspose.Cells convert worksheet data extent to printable range | C# generate A1 address from cell indices Aspose.Cells | dynamic worksheet printing range Aspose.Cells .NET | Excel workbook print region automation with Aspose.Cells | set page setup range based on data extent C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The program loads an existing workbook, retrieves the worksheet's MaxDisplayRange, converts the start and end cells to A1 notation, assigns that address to PageSetup.PrintArea, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or any specific worksheet)
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the used range of the worksheet
                AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                // Calculate start and end cell indices
                int firstRow = usedRange.FirstRow;
                int firstColumn = usedRange.FirstColumn;
                int lastRow = usedRange.RowCount > 0 ? firstRow + usedRange.RowCount - 1 : firstRow;
                int lastColumn = usedRange.ColumnCount > 0 ? firstColumn + usedRange.ColumnCount - 1 : firstColumn;

                // Convert the start and end cells of the used range to A1‑style addresses
                string startCell = CellsHelper.CellIndexToName(firstRow, firstColumn);
                string endCell = CellsHelper.CellIndexToName(lastRow, lastColumn);

                // Set the print area to the used range
                sheet.PageSetup.PrintArea = $"{startCell}:{endCell}";

                // Save the workbook after setting the print area
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed and saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
