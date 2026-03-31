using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsLoadDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] files = {
                "sample.xlsx",
                "sample.csv",
                "sample.ods",
                "sample.pdf"
            };

            foreach (string filePath in files)
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    continue;
                }

                string ext = Path.GetExtension(filePath);
                if (string.Equals(ext, ".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Skipping unsupported format: {filePath}");
                    continue;
                }

                using (Workbook workbook = new Workbook(filePath))
                {
                    Console.WriteLine($"Loaded '{Path.GetFileName(filePath)}' - Worksheets count: {workbook.Worksheets.Count}");

                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue($"Loaded on {DateTime.Now}");

                    string directory = Path.GetDirectoryName(filePath) ?? string.Empty;
                    string outputPath = Path.Combine(
                        directory,
                        $"{Path.GetFileNameWithoutExtension(filePath)}_modified{Path.GetExtension(filePath)}");

                    workbook.Save(outputPath);

                    Console.WriteLine($"Modified workbook saved as: {outputPath}");
                    Console.WriteLine();
                }
            }
        }
    }
}