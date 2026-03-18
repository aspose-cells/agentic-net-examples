using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveUnusedStylesDemo
    {
        public static void Run()
        {
            // Load an existing workbook (replace with your file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Display the number of styles before cleanup
            int beforeCount = workbook.CountOfStylesInPool;
            Console.WriteLine($"Styles in pool before removal: {beforeCount}");

            // Remove all styles that are not applied to any cell
            workbook.RemoveUnusedStyles();

            // Display the number of styles after cleanup
            int afterCount = workbook.CountOfStylesInPool;
            Console.WriteLine($"Styles in pool after removal: {afterCount}");

            // Save the cleaned workbook (replace with your desired output path)
            string outputPath = "output_without_unused_styles.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            RemoveUnusedStylesDemo.Run();
        }
    }
}