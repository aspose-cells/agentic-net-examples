using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VeryHiddenWorksheetDemo
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Access the first worksheet and rename it
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SecretData";

                // Set the worksheet visibility to VeryHidden so it cannot be shown via Excel UI
                sheet.VisibilityType = VisibilityType.VeryHidden;

                // Save the workbook (lifecycle rule: save)
                string outputPath = "VeryHiddenWorksheet.xlsx";
                workbook.Save(outputPath);

                // Verify the file exists before loading
                if (!File.Exists(outputPath))
                {
                    Console.WriteLine($"Error: The file '{outputPath}' was not found.");
                    return;
                }

                // Load the workbook to verify the visibility setting (lifecycle rule: load)
                Workbook loadedWorkbook = new Workbook(outputPath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                Console.WriteLine($"Worksheet '{loadedSheet.Name}' visibility: {loadedSheet.VisibilityType}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}