// Title: Copy Formatting from a Password‑Protected Excel Range (O1:O5 → P1:P5) with Aspose.Cells for .NET
// Description: Load a password‑protected workbook using Aspose.Cells LoadOptions, create the source range O1:O5 and destination range P1:P5, copy only the cell styles with CopyStyle, and save the result to a new file.
// Keywords: Aspose.Cells | C# | .NET | CopyStyle | copy formatting | password protected workbook | LoadOptions | Excel range formatting | cell style transfer
// Common Searches: Aspose.Cells copy only formatting between ranges | load password protected Excel file Aspose.Cells .NET | CopyStyle example for protected workbook | transfer cell style O1:O5 to P1:P5 Aspose.Cells | how to preserve data while copying format in Excel using C#
// Developer Intent: Load a password‑protected workbook and copy only the formatting from O1:O5 to P1:P5.
// Use Cases: Apply the visual style of a secured template column to another column without altering its values. | Automate formatting updates in reports that are distributed with workbook passwords. | Standardize appearance across multiple sheets while keeping sensitive data protected.
// AI Prompts: Generate C# code that opens a password‑protected Excel file with Aspose.Cells and copies only the style from range O1:O5 to P1:P5. | Explain the differences between CopyStyle, Copy, and CopyPasteOptions in Aspose.Cells for .NET. | Provide a step‑by‑step tutorial for loading a protected workbook, copying formatting between two ranges, and saving the file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Load a password‑protected workbook using Aspose.Cells LoadOptions, create the source range O1:O5 and destination range P1:P5, copy only the cell styles with CopyStyle, and save the result to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "myPassword"; // replace with actual password

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the password‑protected workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define source and destination ranges
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("O1:O5");
            Aspose.Cells.Range destinationRange = sheet.Cells.CreateRange("P1:P5");

            // Copy only the formatting (style) from source to destination
            destinationRange.CopyStyle(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
