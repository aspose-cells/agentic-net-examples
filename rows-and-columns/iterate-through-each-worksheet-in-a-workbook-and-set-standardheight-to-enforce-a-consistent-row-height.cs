using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetStandardRowHeightForAllSheets
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Add an extra worksheet for demonstration purposes
                workbook.Worksheets.Add();

                // Desired standard row height in points
                double desiredHeight = 20.0;

                // Set the standard height for each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    sheet.Cells.StandardHeight = desiredHeight;
                }

                // Save the workbook to verify the changes
                string outputPath = "AllSheetsStandardHeight.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            SetStandardRowHeightForAllSheets.Run();
        }
    }
}