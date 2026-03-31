using System;
using System.IO;
using Aspose.Cells;

namespace ExternalConnectionProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (lifecycle rule: use constructor with file name)
            Workbook workbook = new Workbook(sourcePath);

            // Process external links that point to XLS or XLSB files
            foreach (ExternalLink link in workbook.Worksheets.ExternalLinks)
            {
                // Check if the data source ends with .xls or .xlsb (case‑insensitive)
                string dataSource = link.DataSource;
                if (dataSource != null &&
                    (dataSource.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                     dataSource.EndsWith(".xlsb", StringComparison.OrdinalIgnoreCase)))
                {
                    // Example processing: change the extension to .xlsx
                    string newDataSource = Path.ChangeExtension(dataSource, ".xlsx");
                    link.DataSource = newDataSource;
                    Console.WriteLine($"Updated external link: {dataSource} -> {newDataSource}");
                }
            }

            // Save the modified workbook as XLSX (lifecycle rule: use Save(string))
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}