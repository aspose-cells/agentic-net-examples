using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveUnusedStylesDemo
    {
        public static void Run()
        {
            // Load the workbook from an existing file
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Remove all styles that are not used by any cell
            workbook.RemoveUnusedStyles();

            // Save the cleaned workbook to a new file
            string outputPath = "output_cleaned.xlsx";
            workbook.Save(outputPath);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            RemoveUnusedStylesDemo.Run();
        }
    }
}