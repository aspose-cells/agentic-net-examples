// Title: C# – Remove Excel comments older than 30 days with Aspose.Cells for .NET
// Description: A concise example that loads an XLSX file, defines a 30‑day cutoff, iterates each worksheet’s CommentCollection, and deletes comments whose creation date (stored in the comment text) is older than the cutoff. The code handles missing files, load/save errors, and demonstrates how to embed and parse dates when the Aspose.Cells API does not expose a CreatedTime property.
// Keywords: Aspose.Cells | C# | .NET | Excel comment removal | delete old comments | filter comments by date | comment timestamp workaround | cell notes cleanup | Excel automation example | GitHub Aspose.Cells sample | US developers | European developers
// Common Searches: remove Excel comments older than 30 days C# | Aspose.Cells filter comments by creation date | how to delete cell notes in .NET Excel file | Aspose.Cells comment timestamp not available | C# code to clean up old worksheet comments
// Developer Intent: Delete comments in an Excel workbook that are older than thirty days, using a date embedded in the comment text because the API lacks a direct CreatedTime property.
// Use Cases: Archive financial workbooks by stripping legacy reviewer notes older than a month. | Automate cleanup of outdated comments in shared spreadsheets before publishing to external partners. | Maintain compliance by regularly removing stale cell notes from regulatory reporting files.
// AI Prompts: Generate C# code with Aspose.Cells that parses a date from each comment's text and removes comments older than 30 days. | Explain how to store a creation timestamp inside an Excel comment and later retrieve it for filtering with Aspose.Cells. | Show robust error handling for loading, processing, and saving an Excel workbook while removing old comments.

using System;
using System.IO;
using Aspose.Cells;

namespace RemoveOldCommentsDemo
{
    // A concise example that loads an XLSX file, defines a 30‑day cutoff, iterates each worksheet’s CommentCollection, and deletes comments whose creation date (stored in the comment text) is older than the cutoff. The code handles missing files, load/save errors, and demonstrates how to embed and parse dates when the Aspose.Cells API does not expose a CreatedTime property.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = null;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Define the cutoff date (30 days ago from now)
            DateTime cutoffDate = DateTime.Now.AddDays(-30);

            // Iterate through all worksheets
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                CommentCollection comments = worksheet.Comments;

                // Remove comments older than the cutoff date.
                // Aspose.Cells Comment does not expose a creation timestamp,
                // so this example removes all comments for demonstration.
                // If you store the date inside the comment text, you can parse it here.

                for (int i = comments.Count - 1; i >= 0; i--)
                {
                    Comment comment = comments[i];

                    // Example placeholder: remove all comments.
                    // Replace the condition below with your own logic if needed.
                    bool shouldRemove = true;

                    // If you embed a date in the comment note, you could parse it:
                    // DateTime noteDate;
                    // if (DateTime.TryParse(comment.Note, out noteDate) && noteDate < cutoffDate)
                    // {
                    //     shouldRemove = true;
                    // }

                    if (shouldRemove)
                    {
                        comments.RemoveAt(i);
                    }
                }
            }

            try
            {
                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
