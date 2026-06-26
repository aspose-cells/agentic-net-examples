using System;
using System.IO;
using Aspose.Cells;

class RemoveUnusedStylesDemo
{
    static void Main()
    {
        try
        {
            // Path to the workbook that may contain unused styles
            string inputPath = "input_with_unused_styles.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the workbook from the file
            Workbook wb = new Workbook(inputPath);

            // Show the number of styles before cleanup
            Console.WriteLine("Styles before cleanup: " + wb.CountOfStylesInPool);

            // Remove all unused styles from the workbook
            wb.RemoveUnusedStyles();

            // Show the number of styles after cleanup
            Console.WriteLine("Styles after cleanup: " + wb.CountOfStylesInPool);

            // Save the cleaned workbook to a new file
            string outputPath = "cleaned_output.xlsx";
            wb.Save(outputPath);

            Console.WriteLine("Workbook saved after removing unused styles.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}