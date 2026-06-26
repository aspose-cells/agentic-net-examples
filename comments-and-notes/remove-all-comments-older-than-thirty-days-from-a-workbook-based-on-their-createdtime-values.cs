using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCommentCleanup
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    CommentCollection comments = sheet.Comments;

                    // Collect indices of comments to remove
                    List<int> indicesToRemove = new List<int>();

                    foreach (Comment comment in comments)
                    {
                        // Aspose.Cells Comment does not expose creation time,
                        // so this example removes all comments.
                        // If you have custom metadata, apply your own filter here.
                        int idx = comments.IndexOf(comment);
                        if (idx >= 0)
                            indicesToRemove.Add(idx);
                    }

                    // Remove comments starting from the highest index to avoid shifting
                    indicesToRemove.Sort();
                    indicesToRemove.Reverse();
                    foreach (int idx in indicesToRemove)
                    {
                        comments.RemoveAt(idx);
                    }
                }

                // Output workbook path
                string outputPath = "output.xlsx";

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}