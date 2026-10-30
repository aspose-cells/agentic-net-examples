// Title: Create an Excel workbook, insert comments into cells A1 and B2, and save as XLSX with Aspose.Cells for .NET
// Common Searches: how to insert a comment into a specific cell using Aspose.Cells C# | Aspose.Cells C# add multiple cell comments and export to xlsx | save a new workbook to a custom folder with Aspose.Cells .NET | programmatically set comment author and text in Aspose.Cells | create workbook and add comments without opening Excel
// Tags: Aspose.Cells add cell comment C# | Aspose.Cells save workbook as Xlsx C# | Aspose.Cells set comment author C# | Aspose.Cells create workbook programmatically | Aspose.Cells ensure output folder exists

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first worksheet, adds regular comments with authors to cells A1 and B2, ensures the target directory exists, and saves the file as an XLSX workbook using Aspose.Cells for .NET, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add a regular comment to cell A1.
            int commentIndexA1 = sheet.Comments.Add("A1");
            Comment commentA1 = sheet.Comments[commentIndexA1];
            commentA1.Author = "Alice";
            commentA1.Note = "This is a comment on A1.";

            // Add a regular comment to cell B2.
            int commentIndexB2 = sheet.Comments.Add("B2");
            Comment commentB2 = sheet.Comments[commentIndexB2];
            commentB2.Author = "Bob";
            commentB2.Note = "Another comment on B2.";

            // Define output path and ensure its directory exists.
            string outputPath = "ThreadedComments.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook in XLSX format.
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
