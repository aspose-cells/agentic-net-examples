using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Properties;

namespace AsposeCellsExamples
{
    public class SetWorkbookCommentsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access built‑in document properties
                BuiltInDocumentPropertyCollection properties = workbook.BuiltInDocumentProperties;

                // Set multiline comments
                properties.Comments = @"This workbook was generated programmatically.
It contains sample data for demonstration purposes.
Please review the content and provide feedback.";

                // Display comments
                Console.WriteLine("Workbook Comments:");
                Console.WriteLine(properties.Comments);

                // Define output path
                string outputPath = "WorkbookWithComments.xlsx";

                // Save workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorkbookCommentsDemo.Run();
        }
    }
}