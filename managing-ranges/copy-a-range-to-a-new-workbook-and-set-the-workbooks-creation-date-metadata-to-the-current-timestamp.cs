// Title: Copy Excel Range to a New Workbook and Set Creation Date with Aspose.Cells for .NET (C#)
// Description: Load a source workbook, copy a defined cell range (e.g., A1:B5) into a new workbook at a target range (e.g., C1:D5), assign the current timestamp to BuiltInDocumentProperties.CreatedTime, and save the file.
// Keywords: Aspose.Cells copy range | C# copy cells between workbooks | set workbook creation date Aspose | BuiltInDocumentProperties CreatedTime | .NET Excel range transfer | Excel metadata timestamp
// Common Searches: Aspose.Cells copy range to another workbook C# | How to set CreatedTime property in Aspose.Cells | Copy cells A1:B5 to C1:D5 using Aspose.Cells | Update Excel file metadata with current date in .NET
// Developer Intent: Transfer a specific cell block from an existing Excel file to a fresh workbook and record the copy time in the workbook’s creation metadata.
// Use Cases: Generate a standalone report section by extracting a data block from a master file. | Create template files on‑the‑fly while preserving accurate generation timestamps. | Archive selected worksheet ranges with metadata that reflects the exact copy moment.
// AI Prompts: Write C# code with Aspose.Cells that copies a given range from one workbook to another and sets the CreatedTime to DateTime.UtcNow. | Explain error handling for missing source files when copying ranges using Aspose.Cells. | Provide a reusable method: (sourcePath, sourceRange, destPath, destRange) → copy range and update creation date.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Load a source workbook, copy a defined cell range (e.g., A1:B5) into a new workbook at a target range (e.g., C1:D5), assign the current timestamp to BuiltInDocumentProperties.CreatedTime, and save the file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string destinationPath = "CopiedRange.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Define the source range to copy (example: A1:B5)
            AsposeRange sourceRange = sourceWorkbook.Worksheets[0].Cells.CreateRange("A1:B5");

            // Create a new (empty) workbook that will receive the copied range
            Workbook destinationWorkbook = new Workbook();

            // Define the destination range where the data will be placed (example: C1:D5)
            AsposeRange destinationRange = destinationWorkbook.Worksheets[0].Cells.CreateRange("C1:D5");

            // Copy the source range into the destination range
            destinationRange.Copy(sourceRange);

            // Set the workbook's creation date metadata to the current timestamp
            destinationWorkbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // Save the new workbook to disk
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Range copied successfully. File saved as {destinationPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
