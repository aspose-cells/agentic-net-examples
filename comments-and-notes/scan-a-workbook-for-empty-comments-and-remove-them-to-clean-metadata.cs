using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class RemoveEmptyComments
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Load the workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                CommentCollection comments = worksheet.Comments;

                // Iterate backwards to safely remove items
                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = comments[i];

                    // Remove comment if its text is null, empty or whitespace
                    if (string.IsNullOrWhiteSpace(comment.Note))
                    {
                        comments.RemoveAt(comment.Row, comment.Column);
                    }
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}