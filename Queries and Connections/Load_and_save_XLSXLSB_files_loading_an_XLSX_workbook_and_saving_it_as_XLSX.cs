using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the output XLSX file
            string outputPath = "output.xlsx";

            // Load the workbook from the source file (uses Workbook(string) constructor)
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook to the destination file in XLSX format (uses Workbook.Save(string, SaveFormat))
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook loaded from '{0}' and saved as '{1}'.", sourcePath, outputPath);
        }
    }
}