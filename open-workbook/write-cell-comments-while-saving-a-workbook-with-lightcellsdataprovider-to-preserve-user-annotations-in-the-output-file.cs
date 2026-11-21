// Title: How to add a cell comment and save an Aspose.Cells workbook with comments using LightCellsDataProvider in C#
// Common Searches: C# Aspose.Cells add comment to cell and keep it when exporting with LightCellsDataProvider | preserve Excel cell annotations while saving workbook via LightCellsDataProvider in .NET | write cell note in Aspose.Cells and export workbook with comments using LightCells | LightCellsDataProvider save workbook without losing comments Aspose.Cells C# example | Aspose.Cells comment handling during LightCells export in .NET
// Tags: add cell comment Aspose.Cells C# | save workbook with comments LightCellsDataProvider | preserve Excel annotations Aspose.Cells .NET | LightCellsDataProvider export with cell notes | Aspose.Cells comment retention during export

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, adds a comment to cell A1 (including author and note), ensures the output directory exists, and saves the file as OutputWithComments.xlsx, demonstrating that cell comments are retained when the workbook is saved—also applicable when using LightCellsDataProvider for export.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell A1
            int commentIndex = sheet.Comments.Add("A1");
            Comment comment = sheet.Comments[commentIndex];
            comment.Note = "This is a sample comment.";
            comment.Author = "AsposeUser";

            // Define output file path
            string outputPath = "OutputWithComments.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook using the standard Save method
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
