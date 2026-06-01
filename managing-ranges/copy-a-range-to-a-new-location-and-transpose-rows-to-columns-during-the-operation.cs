using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class CopyAndTransposeDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Fill sample data in a horizontal range (A1:E1)
            for (int col = 0; col < 5; col++)
            {
                cells[0, col].PutValue($"Data {col + 1}");
            }

            // Define the source range (first row, 5 columns)
            AsposeRange sourceRange = cells.CreateRange(0, 0, 1, 5);

            // Create PasteOptions and enable transposition
            PasteOptions pasteOptions = new PasteOptions
            {
                PasteType = PasteType.All, // copy everything (values, formats, etc.)
                Transpose = true           // transpose rows ↔ columns during paste
            };

            // Define the destination range where the transposed data will be placed
            // Since source is 1x5, transposed will be 5x1, so we create a vertical range starting at A3
            AsposeRange destRange = cells.CreateRange(2, 0, 5, 1);

            // Copy the source range to the destination range with transposition
            destRange.Copy(sourceRange, pasteOptions);

            // Save the workbook to a file
            string outputPath = "CopyAndTransposeDemo.xlsx";
            workbook.Save(outputPath);
        }
    }
}