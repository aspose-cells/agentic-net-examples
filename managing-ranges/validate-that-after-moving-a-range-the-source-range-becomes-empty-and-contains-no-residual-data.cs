using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeMoveValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the source range (A1:B2) and fill it with sample data
                AsposeRange sourceRange = cells.CreateRange("A1:B2");
                sourceRange[0, 0].PutValue("R1C1");
                sourceRange[0, 1].PutValue("R1C2");
                sourceRange[1, 0].PutValue("R2C1");
                sourceRange[1, 1].PutValue("R2C2");

                // Verify that the source range is not blank before moving
                Console.WriteLine("Source range blank before move: " + sourceRange.IsBlank());

                // Move the range to a new location (C3:D4). Row and column indices are zero‑based.
                // Destination start row = 2 (C3), destination start column = 2 (C)
                sourceRange.MoveTo(2, 2);

                // After moving, create a range object that points to the original location
                AsposeRange originalLocation = cells.CreateRange("A1:B2");

                // Validate that the original location is now empty
                bool isEmpty = originalLocation.IsBlank();
                Console.WriteLine("Source range empty after move: " + isEmpty);

                // Save the workbook (optional, just to visualize the result)
                string outputPath = "RangeMoveValidation.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}