using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart, ChartType

namespace AsposeCellsDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Banana");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(30);
                worksheet.Cells["B3"].PutValue(45);

                // 2. Add a column chart that uses the data above
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B3", true);          // Values
                chart.NSeries.CategoryData = "A2:A3";      // Categories

                // 3. Save the workbook to a memory stream in XLSX format
                using MemoryStream stream = new MemoryStream();
                workbook.Save(stream, SaveFormat.Xlsx);
                stream.Position = 0; // Reset position before further processing

                // 4. Save the stream to a local file
                string outputPath = "chartWorkbook.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write the file safely
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await stream.CopyToAsync(file);
                }

                Console.WriteLine($"Workbook with chart saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}