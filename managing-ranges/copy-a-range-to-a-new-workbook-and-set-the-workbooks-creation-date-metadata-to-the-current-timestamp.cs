// Title: Copy a specific cell range to a new workbook and set the workbook's creation date metadata with Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing Excel file, copy the range A1:C10 into a freshly created workbook, and assign DateTime.Now to the new workbook's BuiltInDocumentProperties.CreatedTime. | Create a new workbook, transfer a defined cell block from a source worksheet, then update the workbook's creation timestamp before saving the file.
// Common Searches: aspocells c# copy range A1:C10 to new workbook and set createdtime | how to set workbook creation date property after copying cells with Aspose.Cells | copy cell block from one Excel file to another using Aspose.Cells .NET | Aspose.Cells set BuiltInDocumentProperties.CreatedTime to current time in C# | example of copying a range and updating metadata in Aspose.Cells for .NET
// Tags: Aspose.Cells copy range to new workbook | BuiltInDocumentProperties CreatedTime C# | copy cell block A1:C10 Aspose.Cells | set workbook metadata timestamp Aspose.Cells | C# Aspose.Cells range.Copy example

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// Copies the A1:C10 range from a source workbook into a new workbook, updates the new workbook's CreatedTime property to the current timestamp, and saves the result.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "SourceWorkbook.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook srcWorkbook = new Workbook(sourcePath);

            // Create a new workbook (contains a default worksheet)
            Workbook newWorkbook = new Workbook();

            // Define the range to copy from the source workbook (e.g., A1:C10 on the first worksheet)
            AsposeRange srcRange = srcWorkbook.Worksheets[0].Cells.CreateRange("A1:C10");

            // Define the destination range in the new workbook (starting at A1 on the first worksheet)
            AsposeRange destRange = newWorkbook.Worksheets[0].Cells.CreateRange("A1");

            // Copy the source range to the destination range
            srcRange.Copy(destRange);

            // Set the workbook's creation date metadata to the current timestamp
            newWorkbook.BuiltInDocumentProperties.CreatedTime = DateTime.Now;

            // Path for the output workbook
            string outputPath = "CopiedRangeWorkbook.xlsx";

            // Save the new workbook
            newWorkbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
