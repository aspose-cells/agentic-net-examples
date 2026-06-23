using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ClearWorksheetBackgroundImage
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one if needed)
            Workbook workbook = new Workbook();

            // Get the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Clear the worksheet's background image by setting the property to null
            worksheet.BackgroundImage = null;

            // Define output file path
            string outputPath = "WorksheetWithoutBackground.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to verify that the background image has been removed
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}