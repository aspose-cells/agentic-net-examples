using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableDownlevelRevealedCommentsDemo
    {
        // Entry point for the application
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook (in-memory)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add a sample comment to demonstrate the effect
                int commentIndex = sheet.Comments.Add("A1");
                Comment comment = sheet.Comments[commentIndex];
                comment.Note = "Sample comment";

                // Create HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Disable downlevel-revealed conditional comments in the HTML output
                    DisableDownlevelRevealedComments = true
                };

                // Define output file path
                string outputPath = "Output_NoDownlevelComments.html";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine("Workbook saved with downlevel-revealed comments disabled.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}