// Title: Set all worksheets to Normal view mode in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing XLSX file with Aspose.Cells, loops through every worksheet, assigns ViewType.Normal (using a safe cast), and saves the workbook to a new location, including checks for missing input files and creating the output directory. | Generate a C# snippet that changes the view mode of each sheet in a workbook to Normal via Aspose.Cells, demonstrates proper enum casting, and persists the changes with Workbook.Save while handling possible I/O errors.
// Common Searches: aspocells c# set worksheet viewtype to normal for all sheets | how to change Excel sheet view to normal using Aspose.Cells .NET | C# loop through workbook worksheets and set view mode to normal with Aspose.Cells | save workbook after modifying view settings with Aspose.Cells in C#
// Tags: Aspose.Cells set worksheet viewtype normal | C# iterate workbook worksheets Aspose.Cells | Excel normal view mode Aspose.Cells API | Workbook.Save after viewtype change | handle missing input file Aspose.Cells C#

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing XLSX file, iterates over each worksheet to set its ViewType to Normal (using a numeric cast for compatibility), ensures the output directory exists, and saves the modified workbook while providing robust error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Set each worksheet to Normal view mode
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use numeric cast to avoid enum member name issues in certain library versions
                sheet.ViewType = (ViewType)0; // Normal view
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
