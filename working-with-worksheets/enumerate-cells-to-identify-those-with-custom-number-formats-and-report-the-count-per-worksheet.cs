using System;
using Aspose.Cells;

namespace AsposeCellsCustomNumberFormatCounter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Counter for cells with custom number formats in the current worksheet
                int customFormatCount = 0;

                // Get the Cells collection of the worksheet
                Cells cells = sheet.Cells;

                // Iterate over all instantiated cells (efficient for sparse data)
                foreach (Cell cell in cells)
                {
                    // Retrieve the style of the cell
                    Style style = cell.GetStyle();

                    // If the Custom property is not empty, a custom number format is applied
                    if (!string.IsNullOrEmpty(style.Custom))
                    {
                        customFormatCount++;
                    }
                }

                // Output the count for the current worksheet
                Console.WriteLine($"Worksheet \"{sheet.Name}\": {customFormatCount} cell(s) with custom number formats.");
            }

            // Optionally, save the workbook if any modifications were made
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}