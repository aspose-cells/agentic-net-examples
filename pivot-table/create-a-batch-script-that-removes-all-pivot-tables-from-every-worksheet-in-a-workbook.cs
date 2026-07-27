using System;
using Aspose.Cells;

namespace RemoveAllPivotTables
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook (replace with your actual file path)
            string inputPath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Remove all pivot tables from every worksheet in the workbook
            workbook.Worksheets.ClearPivottables();

            // Path to save the modified workbook (replace with desired output path)
            string outputPath = "output_without_pivots.xlsx";

            // Save the workbook after clearing pivot tables
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("All pivot tables have been removed and the workbook saved to: " + outputPath);
        }
    }
}