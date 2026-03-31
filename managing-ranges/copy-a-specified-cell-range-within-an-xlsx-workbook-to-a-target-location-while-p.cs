using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeCopyDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (you can change the index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define the source range to copy (e.g., B2:D5)
            AsposeRange sourceRange = cells.CreateRange("B2:D5");

            // Define the destination range where the source will be copied (e.g., F2:H5)
            // The destination range must have the same size as the source range.
            AsposeRange destinationRange = cells.CreateRange("F2:H5");

            // Set paste options to copy everything (values, formulas, formatting, etc.)
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All,   // Preserve all content
                SkipBlanks = false,          // Do not skip blanks
                Transpose = false            // No transposition
            };

            // Perform the copy operation using the destination range's Copy method
            destinationRange.Copy(sourceRange, pasteOptions);

            // Save the workbook with the copied range
            workbook.Save("output.xlsx");
        }
    }
}