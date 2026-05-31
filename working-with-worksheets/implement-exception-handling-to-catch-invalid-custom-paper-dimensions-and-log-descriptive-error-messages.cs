using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomPaperSizeDemo
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

                // Example custom dimensions (in inches)
                double widthInches = 8.5;
                double heightInches = -11.0; // Invalid height to demonstrate error handling

                // Convert inches to points (1 inch = 72 points) as required by Aspose.Cells
                double widthPoints = widthInches * 72;
                double heightPoints = heightInches * 72;

                try
                {
                    // Attempt to set a custom paper size
                    worksheet.PageSetup.CustomPaperSize(widthPoints, heightPoints);
                    Console.WriteLine($"Custom paper size set: Width={widthInches} inches, Height={heightInches} inches.");
                }
                catch (CellsException ex)
                {
                    // Handle page‑setup related errors
                    Console.WriteLine($"Error: Invalid custom paper dimensions (Width={widthInches}, Height={heightInches}).");
                    Console.WriteLine($"Exception Message: {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Handle any other unexpected errors
                    Console.WriteLine($"Unexpected error while setting paper size: {ex.Message}");
                }

                // Save the workbook (will reflect the page setup if valid)
                string outputPath = "CustomPaperSizeResult.xlsx";

                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine("Workbook saved successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // Global safety net
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }
    }
}