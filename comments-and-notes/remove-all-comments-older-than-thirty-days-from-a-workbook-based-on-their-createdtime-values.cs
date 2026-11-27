// Title: Delete Excel worksheet comments with Aspose.Cells for .NET – iterate worksheets and remove comments
// Common Searches: c# code to delete all comments from an Excel workbook using Aspose.Cells | how to remove Excel cell notes programmatically with Aspose.Cells in .NET | asp.net iterate worksheet comments collection and delete them | filter and delete Excel comments older than a specific date using Aspose.Cells | remove Excel comments in bulk with Aspose.Cells C# example | Aspose.Cells comment removal loop backwards C#
// Tags: Aspose.Cells delete worksheet comments | C# remove Excel cell notes | iterate comments collection Aspose.Cells | bulk comment removal .NET Excel | filter comments by date Aspose.Cells | backwards loop comment deletion C#

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file with Aspose.Cells, iterates each worksheet, and removes every comment by looping backwards through the Comments collection (the API does not expose a creation timestamp, so custom date logic would be needed for age‑based filtering). The modified workbook is then saved.
class RemoveOldComments
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var comments = sheet.Comments;

                // Loop backwards to safely remove comments while iterating
                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    // Aspose.Cells.Comment does not expose a creation date.
                    // Remove the comment (or add custom date logic here if needed).
                    comments.RemoveAt(i);
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
