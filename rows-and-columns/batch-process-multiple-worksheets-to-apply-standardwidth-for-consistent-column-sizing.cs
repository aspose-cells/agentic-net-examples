using System;
using Aspose.Cells;

namespace AsposeCellsBatchStandardWidth
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // For loading: new Workbook("input.xlsx");

            // Add sample worksheets (optional, for demonstration)
            workbook.Worksheets.Add();
            workbook.Worksheets.Add();

            // Desired standard column width (in characters)
            double standardWidth = 18.25;

            // Apply the standard width to every worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells.StandardWidth = standardWidth;
                // Verify the assignment (optional)
                Console.WriteLine($"Worksheet '{sheet.Name}' StandardWidth set to {sheet.Cells.StandardWidth}");
            }

            // Save the modified workbook
            workbook.Save("BatchStandardWidth.xlsx");
        }
    }
}

// Author: Example demonstrating batch application of Cells.StandardWidth across all worksheets.