using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveEmptyComments
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate backwards through comments to safely remove items
                for (int i = worksheet.Comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = worksheet.Comments[i];

                    // Remove comment if its text is null, empty, or whitespace
                    if (string.IsNullOrWhiteSpace(comment.Note))
                    {
                        worksheet.Comments.RemoveAt(comment.Row, comment.Column);
                    }
                }
            }

            // Save the cleaned workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}