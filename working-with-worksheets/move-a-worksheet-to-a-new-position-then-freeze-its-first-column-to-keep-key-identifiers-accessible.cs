using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MoveAndFreezeWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet (you can choose any worksheet)
                Worksheet worksheet = workbook.Worksheets[0];

                // Move the worksheet to a new position (e.g., index 2 -> third tab)
                worksheet.MoveTo(2);

                // Freeze the first column (no rows frozen, first column frozen)
                worksheet.FreezePanes(0, 1, 0, 1);

                // Define output file path
                string outputPath = "MovedAndFrozen.xlsx";

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MoveAndFreezeWorksheet.Run();
        }
    }
}