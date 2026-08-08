// Title: C# – Delete Blank Comments in Excel with Aspose.Cells
// Description: Loads an .xlsx file with Aspose.Cells, scans every worksheet’s CommentCollection, and eliminates comments whose Note is null, empty or only whitespace, then writes the cleaned workbook to a new file.
// Keywords: Aspose.Cells C# remove blank comments | delete empty Excel notes .NET | clean workbook comment metadata | iterate worksheet comments Aspose | strip whitespace comments C# | Excel comment cleanup Aspose.Cells
// Common Searches: How to delete blank comments in an Excel file using Aspose.Cells C# | Remove empty notes from all sheets with Aspose.Cells .NET | Aspose.Cells example for cleaning comment collection | C# code to purge whitespace‑only comments from a workbook
// Developer Intent: Remove every comment that contains no visible text from all worksheets in an Excel workbook.
// Use Cases: Prepare a report for distribution by stripping placeholder comments. | Minimize file size and improve performance by discarding unnecessary comment metadata. | Ensure data migration scripts only transfer meaningful annotations.
// AI Prompts: Generate C# code that uses Aspose.Cells to iterate through each worksheet and delete comments whose Note property is null, empty, or whitespace, then save the workbook. | Show how to safely remove empty comments from a CommentCollection by looping backwards to avoid index errors.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads an .xlsx file with Aspose.Cells, scans every worksheet’s CommentCollection, and eliminates comments whose Note is null, empty or only whitespace, then writes the cleaned workbook to a new file.
    public class RemoveEmptyComments
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
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Get the comments collection of the current worksheet
                CommentCollection comments = worksheet.Comments;

                // Iterate backwards so that removal does not affect the loop index
                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = comments[i];

                    // Remove comment if its text is null, empty, or whitespace
                    if (string.IsNullOrWhiteSpace(comment.Note))
                    {
                        comments.RemoveAt(i);
                    }
                }
            }

            // Save the cleaned workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
