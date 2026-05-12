using System;
using System.Drawing;
using Aspose.Cells;

namespace CommentFontColorModifier
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Path for the modified workbook
            string outputPath = "output.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Iterate through all comments in the current worksheet
                foreach (Comment comment in worksheet.Comments)
                {
                    // Change the comment font color based on the author name
                    // (example logic – adjust as needed)
                    if (comment.Author == "Alice")
                    {
                        comment.Font.Color = Color.Red;
                    }
                    else if (comment.Author == "Bob")
                    {
                        comment.Font.Color = Color.Blue;
                    }
                    else
                    {
                        comment.Font.Color = Color.Green;
                    }
                }
            }

            // Save the modified workbook (lifecycle rule: save)
            workbook.Save(outputPath);

            // Optional: release resources
            workbook.Dispose();

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}