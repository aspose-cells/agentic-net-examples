// Title: Insert a signer name and title as a comment‑based signature line on a specific worksheet using Aspose.Cells for .NET
// Common Searches: C# add a comment to cell A1 as a signature line with Aspose.Cells | Aspose.Cells set signer name and title in an Excel worksheet | how to simulate a signature line in .xlsx using Aspose.Cells .NET | add worksheet annotation for signer details with Aspose.Cells C# | insert comment containing signature information into a specific sheet using Aspose.Cells
// Tags: add comment to worksheet cell Aspose.Cells | set comment text signer name title C# | simulate signature line Excel Aspose.Cells | worksheet annotation insertion Aspose.Cells .NET | save workbook after comment Aspose.Cells | comment-based signature implementation C# | Excel cell A1 comment Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing .xlsx file, accesses a chosen worksheet, creates a comment at cell A1 containing the signer’s name and title to act as a signature line, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the specific worksheet (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name or use index

            // Add a comment to represent a signature line (Aspose.Cells does not have a SignatureLine class)
            // The comment contains signer name and title.
            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "Signed by John Doe, Chief Executive Officer";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
