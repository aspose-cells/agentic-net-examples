using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeReplace
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook
            string inputPath = "InputWorkbook.xlsx";

            // Path to the destination workbook
            string outputPath = "OutputWorkbook.xlsx";

            // The name of the existing named range to be replaced
            string oldRangeName = "OldRange";

            // The new range address (e.g., "B2:C5") that will replace the old range
            string newRangeAddress = "B2:C5";

            // Load the workbook from file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range object by its name
            Name namedRange = workbook.Worksheets.Names[oldRangeName];

            if (namedRange != null)
            {
                // Determine the worksheet that the named range belongs to.
                // If the name is global (SheetIndex == 0), use the first worksheet.
                int sheetIndex = namedRange.SheetIndex > 0 ? namedRange.SheetIndex - 1 : 0;
                Worksheet targetSheet = workbook.Worksheets[sheetIndex];

                // Update the RefersTo property to point to the new range.
                // The RefersTo string must start with an equal sign and include the sheet name.
                namedRange.RefersTo = $"={targetSheet.Name}!{newRangeAddress}";
            }
            else
            {
                Console.WriteLine($"Named range '{oldRangeName}' not found.");
                return;
            }

            // Recalculate all formulas in the workbook to reflect the changed range.
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}