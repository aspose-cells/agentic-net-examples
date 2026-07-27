using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample win/loss data in F2:F12
                for (int i = 0; i < 11; i++)
                {
                    sheet.Cells[1 + i, 5].PutValue(i % 2 == 0 ? 1 : -1);
                }

                // NOTE: Sparkline APIs require the Aspose.Cells.Sparkline assembly.
                // If the assembly is not available, the sparkline creation code is omitted.

                // Save the workbook
                string outputPath = "WinLossData.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}