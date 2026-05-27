using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class PrintCommentsInPlaceDemo
    {
        public static void Main()
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
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data to cells (optional)
            sheet.Cells["A1"].PutValue("First cell");
            sheet.Cells["B2"].PutValue("Second cell");

            // Add comments to the cells
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "Comment for A1";

            commentIndex = sheet.Comments.Add("B2");
            comment = sheet.Comments[commentIndex];
            comment.Note = "Comment for B2";

            // Set the PrintComments property to print comments in place
            sheet.PageSetup.PrintComments = PrintCommentsType.PrintInPlace;

            // Define output file path
            string outputPath = "PrintCommentsInPlace.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}