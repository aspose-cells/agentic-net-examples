// Title: Copy a worksheet from one Excel file to another in C# with Aspose.Cells without creating extra Workbook objects
// AI Prompts: Write C# code that copies the 'Sheet1' worksheet from source.xlsx to a new workbook and saves it as destination.xlsx using Aspose.Cells. | Generate a C# example that copies all worksheets from a source workbook to a destination workbook, preserving original sheet names and removing the default sheet created by Aspose.Cells. | Provide C# error‑handling logic to check for the existence of the source Excel file before performing a worksheet copy with Aspose.Cells.
// Common Searches: aspocells copy worksheet to new workbook c# without intermediate workbook | c# aspocells addcopy method example | how to copy sheet from one Excel file to another using aspocells | aspocells copy multiple sheets preserving names c# | remove default sheet after addcopy aspocells c#
// Tags: aspocells copy worksheet between workbooks c# | addcopy method aspocells c# | copy sheet without intermediate workbook aspocells | remove default sheet after addcopy aspocells | validate source file existence aspocells c#

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the source file, loads it into a Workbook, creates a new destination Workbook, uses the AddCopy method to duplicate the specified worksheet, optionally removes the automatically created default sheet, and saves the result as a new Excel file, with comprehensive exception handling.
class LightCellsCopyExample
{
    static void Main()
    {
        // Paths to the source and destination Excel files
        string sourceFilePath = "source.xlsx";
        string destinationFilePath = "destination.xlsx";

        try
        {
            // Verify source file exists
            if (!File.Exists(sourceFilePath))
                throw new FileNotFoundException($"Source file not found: {sourceFilePath}");

            // Ensure destination directory exists (if a directory is specified)
            string destDir = Path.GetDirectoryName(destinationFilePath);
            if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourceFilePath);

            // Create a new workbook for the destination
            Workbook destinationWorkbook = new Workbook();

            // Get the worksheet to copy (assumed name "Sheet1")
            Worksheet sourceSheet = sourceWorkbook.Worksheets["Sheet1"];
            if (sourceSheet == null)
                throw new InvalidOperationException("Worksheet 'Sheet1' not found in source file.");

            // Add a copy of the source worksheet to the destination workbook
            // AddCopy expects the worksheet name, not the Worksheet object
            destinationWorkbook.Worksheets.AddCopy(sourceSheet.Name);

            // Optionally remove the default empty sheet created with a new workbook
            if (destinationWorkbook.Worksheets.Count > 1 && destinationWorkbook.Worksheets[0].Name == "Sheet1")
                destinationWorkbook.Worksheets.RemoveAt(0);

            // Save the destination workbook
            destinationWorkbook.Save(destinationFilePath);

            Console.WriteLine("Data copied successfully from source.xlsx to destination.xlsx.");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
