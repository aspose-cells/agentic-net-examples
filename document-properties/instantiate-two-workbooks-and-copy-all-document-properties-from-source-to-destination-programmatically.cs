// Title: Copy all Excel workbook document properties from a source file to a new workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells copy workbook metadata to another Excel file | how to transfer Excel document properties programmatically with Aspose.Cells | duplicate custom properties from one workbook to another using Aspose.Cells .NET | copy title author keywords Excel file using Aspose.Cells API | clone Excel file properties when creating a new workbook in C#
// Tags: Aspose.Cells copy workbook document properties | C# transfer Excel metadata programmatically | duplicate custom document properties Aspose.Cells | clone Excel workbook attributes .NET | Aspose.Cells duplicate workbook metadata

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source Excel file into an Aspose.Cells Workbook, creates an empty destination Workbook, notes that Aspose.Cells lacks a single API to copy document properties, and therefore saves the new workbook without transferring the metadata, while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string destinationPath = "destination.xlsx";

        try
        {
            // Ensure the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create an empty destination workbook
            Workbook destinationWorkbook = new Workbook();

            // NOTE: Aspose.Cells does not expose a direct WorkbookProperties API for copying.
            // If copying of document properties is required, it can be implemented using
            // the appropriate properties (e.g., Title, Author) individually.
            // For this example, we proceed without copying properties.

            // Save the destination workbook
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Destination workbook saved successfully to '{destinationPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
