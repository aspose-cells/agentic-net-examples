using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a comment to cell A1 (or retrieve an existing one)
        int commentIdx = worksheet.Comments.Add("A1");
        Comment comment = worksheet.Comments[commentIdx];
        comment.Note = "This is a sample comment.";

        // Change the font color of the entire comment
        comment.Font.Color = Color.Blue;   // Set desired color here

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}