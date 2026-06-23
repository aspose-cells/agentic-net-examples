using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class DateSequenceAutoFillDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Initialize the first two dates (Jan 1, 2023 and Jan 2, 2023)
            cells["A1"].PutValue(new DateTime(2023, 1, 1));
            cells["A2"].PutValue(new DateTime(2023, 1, 2));

            // Define the source range (the two dates) and the target range to fill
            AsposeRange sourceRange = cells.CreateRange("A1:A2");
            AsposeRange targetRange = cells.CreateRange("A3:A10");

            // Use AutoFill with the Series type to extend the date sequence
            sourceRange.AutoFill(targetRange, AutoFillType.Series);

            // Apply a consistent date number format (e.g., "mm-dd-yyyy") to column A
            for (int row = 0; row <= 9; row++)
            {
                Cell cell = cells[row, 0]; // Column A
                Style style = cell.GetStyle();
                style.Custom = "mm-dd-yyyy"; // Custom date format
                style.Number = 14;           // Built‑in date format (optional)
                cell.SetStyle(style);
            }

            // Save the workbook
            string outputPath = "DateSequenceAutoFill.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}