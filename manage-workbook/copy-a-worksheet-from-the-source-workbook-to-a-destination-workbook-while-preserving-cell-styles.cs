// Title: Copy a worksheet from one Excel file to another while preserving all cell styles using Aspose.Cells for .NET
// AI Prompts: Load source.xlsx with Aspose.Cells, create an empty workbook, remove its default sheet, and copy the first worksheet preserving formatting to destination.xlsx. | Generate C# code that uses Worksheets.AddCopy to duplicate a worksheet by name, ensures the target folder exists, and saves the new workbook while keeping all styles intact.
// Common Searches: Aspose.Cells C# copy worksheet to new workbook keep formatting | How to preserve cell styles when copying an Excel sheet with Aspose.Cells | Copy first sheet from source.xlsx to destination.xlsx using Aspose.Cells .NET | Remove default worksheet before adding copied sheet Aspose.Cells | Save copied worksheet to a different file path in C# Aspose.Cells
// Tags: worksheets.addcopy method preserve formatting | remove default worksheet Aspose.Cells | save workbook to specified path C# | verify source file existence Aspose.Cells | create empty workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks that source.xlsx exists, loads it, creates a new workbook without the default sheet, copies the first worksheet by name using Worksheets.AddCopy (which retains all cell styles), ensures the destination directory is present, and saves the result as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destPath = "destination.xlsx";

            // Verify source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create destination workbook and remove default sheet
            Workbook destinationWorkbook = new Workbook();
            destinationWorkbook.Worksheets.Clear();

            // Get the first worksheet from source
            Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];

            // Copy the worksheet by name to the destination workbook
            destinationWorkbook.Worksheets.AddCopy(sourceWorksheet.Name);

            // Ensure the destination directory exists
            string destDirectory = Path.GetDirectoryName(destPath);
            if (!string.IsNullOrEmpty(destDirectory) && !Directory.Exists(destDirectory))
            {
                Directory.CreateDirectory(destDirectory);
            }

            // Save the destination workbook
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Worksheet copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
