// Title: Insert a mandatory signature comment into a specific Excel cell using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds a visible comment to cell B6 as a required signature placeholder and saves the workbook. | Modify the example to accept a cell address parameter and enforce that the comment text is not empty before saving. | Add validation that throws an exception if the comment author or note is missing, and ensure the output folder is created automatically.
// Common Searches: how to add a required signature comment to a cell with Aspose.Cells C# | Aspose.Cells create visible comment for signer in Excel workbook | C# ensure output directory exists when saving Excel file with Aspose.Cells | validate comment author and note before saving workbook using Aspose.Cells | add mandatory comment placeholder for signature line in .xlsx using .NET
// Tags: add visible comment Aspose.Cells C# | mandatory signature placeholder Excel Aspose.Cells | create comment in specific cell Aspose.Cells | ensure output directory exists Aspose.Cells save | validate comment fields before workbook save Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates creating a new workbook with Aspose.Cells, inserting a visible comment in cell B6 as a mandatory signature placeholder, ensuring the target directory exists, and saving the file as SignatureLineWithMandatoryComment.xlsx.
class AddSignatureLineWithMandatoryComment
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a comment to cell B6 (row index 5, column index 2) as a placeholder for a signature line.
            int row = 5;
            int column = 2;
            Cell targetCell = sheet.Cells[row, column];
            int commentIndex = sheet.Comments.Add(row, column);
            Comment comment = sheet.Comments[commentIndex];
            comment.Author = "John Doe";
            comment.Note = "Please sign and add your comments.";
            comment.IsVisible = true; // Make the comment visible

            // Save the workbook to a file (lifecycle rule: save)
            string outputPath = "SignatureLineWithMandatoryComment.xlsx";

            // Ensure the directory exists before saving (handle cases where outputPath has no directory part)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
