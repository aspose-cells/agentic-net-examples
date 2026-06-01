using System;
using System.IO;
using Aspose.Cells;

namespace RemoveOldCommentsDemo
{
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    CommentCollection comments = worksheet.Comments;

                    // Iterate backwards so that removal does not affect the loop index
                    for (int i = comments.Count - 1; i >= 0; i--)
                    {
                        // Remove the comment at the current index
                        comments.RemoveAt(i);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}