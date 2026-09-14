// Title: Copy only formatting from one range to another in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Copy the formatting of cells A1:C3 to cells E1:G3 in a workbook while keeping the original values unchanged using Aspose.Cells C#. | Use Aspose.Cells PasteOptions with PasteType.Formats to transfer styles between two equal‑sized ranges in a .NET application.
// Common Searches: Aspose.Cells how to copy cell styles without values in C# | Copy only formatting between ranges in Excel using Aspose.Cells .NET | PasteOptions PasteType.Formats example for range formatting transfer | Preserve cell values while copying formatting with Aspose.Cells | Copy formatting from A1:C3 to E1:G3 using Aspose.Cells API
// Tags: range formatting copy Aspose.Cells | PasteOptions Formats C# | copy cell styles Aspose.Cells | preserve cell values while copying formatting .NET | Excel range formatting transfer Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using AsposeRange = Aspose.Cells.Range;

// The program loads source.xlsx, defines source range A1:C3 and destination range E1:G3 on the first worksheet, copies only the formatting using PasteOptions with PasteType.Formats, and saves the result to output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "source.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source range (A1:C3)
            AsposeRange sourceRange = sheet.Cells.CreateRange("A1", "C3");

            // Define the destination range (E1:G3) – must be the same size as the source
            AsposeRange destRange = sheet.Cells.CreateRange("E1", "G3");

            // Set up paste options to transfer only formatting
            PasteOptions pasteOptions = new PasteOptions
            {
                // Copy only cell formats (includes conditional formats in most versions)
                PasteType = PasteType.Formats
            };

            // Perform the copy operation using the defined options
            destRange.Copy(sourceRange, pasteOptions);

            // Save the workbook with the applied formatting
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
