// Title: Copy only formatting from range O1:O5 to P1:P5 in a password‑protected Excel workbook using Aspose.Cells for .NET
// AI Prompts: Open a password‑protected .xlsx file with Aspose.Cells, copy the style from cells O1‑O5 to P1‑P5, and save the updated workbook. | Using Aspose.Cells for .NET, load a workbook with a password, transfer only the formatting between two column ranges, and write the result to a new file.
// Common Searches: Aspose.Cells C# copy cell style from one range to another in a protected workbook | How to load a password protected Excel file and copy only formatting using Aspose.Cells | Copy formatting O1 to O5 into P1 to P5 with Aspose.Cells .NET example | C# Aspose.Cells LoadOptions password copy style between ranges | Transfer column formatting in an encrypted Excel file using Aspose.Cells
// Tags: Aspose.Cells Range.CopyStyle usage | load password protected Excel workbook C# | copy cell formatting between columns Aspose.Cells | C# workbook.Save after style transfer | Aspose.Cells protected file formatting example

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Loads a password‑protected workbook, copies only the formatting from cells O1‑O5 to P1‑P5 on the first worksheet, and saves the modified file as output.xlsx.
class Program
{
    static void Main()
    {
        const string inputPath = "protected.xlsx";
        const string outputPath = "output.xlsx";
        const string password = "yourPassword"; // replace with actual password

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
            AsposeRange sourceRange = sheet.Cells.CreateRange("O1", "O5");
            AsposeRange destinationRange = sheet.Cells.CreateRange("P1", "P5");

            // Copy only the formatting (style) from source to destination
            sourceRange.CopyStyle(destinationRange);

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
