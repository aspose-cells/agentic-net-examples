using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class AddNewCommentDemo
    {
        public static void Run()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a new comment to cell A1
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "New comment added.";
            comment.Author = "User";

            // Save the workbook with the new comment
            workbook.Save("output.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            AddNewCommentDemo.Run();
        }
    }
}