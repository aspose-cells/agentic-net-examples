using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetWorksheetDefaultPercentageFormat
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Get the workbook's default style
                Style defaultStyle = workbook.DefaultStyle;

                // Set the built‑in number format to percentage with two decimal places (0.00%)
                // Built‑in format index 10 corresponds to "0.00%" as per Aspose.Cells documentation
                defaultStyle.Number = 10;

                // Save the workbook (lifecycle rule: save)
                string outputPath = "WorksheetDefaultPercentageFormat.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point required for console application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetWorksheetDefaultPercentageFormat.Run();
        }
    }
}