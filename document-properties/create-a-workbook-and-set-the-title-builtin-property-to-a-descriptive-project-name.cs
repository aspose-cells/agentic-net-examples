using System;
using Aspose.Cells;
using System.IO;

namespace AsposeCellsExamples
{
    public class SetWorkbookTitleDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook workbook = new Workbook();

                // Set the built‑in Title property to a descriptive project name
                workbook.BuiltInDocumentProperties.Title = "Project XYZ - Financial Report";

                // Define output file path
                string outputPath = "ProjectReport.xlsx";

                // Save the workbook to a file (lifecycle: save)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorkbookTitleDemo.Run();
        }
    }
}