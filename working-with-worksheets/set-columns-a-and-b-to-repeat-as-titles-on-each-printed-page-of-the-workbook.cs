using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RepeatColumnsAsPrintTitles
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data
                for (int row = 0; row < 100; row++)
                {
                    worksheet.Cells[row, 0].PutValue($"Title A - Row {row + 1}");
                    worksheet.Cells[row, 1].PutValue($"Title B - Row {row + 1}");
                    worksheet.Cells[row, 2].PutValue($"Data C - Row {row + 1}");
                    worksheet.Cells[row, 3].PutValue($"Data D - Row {row + 1}");
                }

                // Set columns A and B to repeat on each printed page
                worksheet.PageSetup.PrintTitleColumns = "$A:$B";

                // Save the workbook
                string outputPath = "RepeatColumnsPrintTitles.xlsx";

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            RepeatColumnsAsPrintTitles.Run();
        }
    }
}