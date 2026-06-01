using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class SheetPaperSizeCollector
    {
        static void Main()
        {
            // Create or load a workbook (replace with actual file path if needed)
            Workbook workbook = new Workbook(); // new workbook; for loading use new Workbook("input.xlsx");

            // Dictionary to hold sheet name and its paper dimensions (width, height) in inches
            Dictionary<string, Tuple<double, double>> sheetPaperSizes = new Dictionary<string, Tuple<double, double>>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access the PageSetup of the current worksheet
                PageSetup pageSetup = sheet.PageSetup;

                // Retrieve paper width and height (in inches)
                double width = pageSetup.PaperWidth;
                double height = pageSetup.PaperHeight;

                // Store the dimensions using the sheet name as the key
                sheetPaperSizes[sheet.Name] = Tuple.Create(width, height);
            }

            // Example: output the collected values to the console
            foreach (var entry in sheetPaperSizes)
            {
                Console.WriteLine($"Sheet: {entry.Key}, Paper Width: {entry.Value.Item1} in, Paper Height: {entry.Value.Item2} in");
            }

            // Save the workbook if any modifications were made (optional)
            // workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}