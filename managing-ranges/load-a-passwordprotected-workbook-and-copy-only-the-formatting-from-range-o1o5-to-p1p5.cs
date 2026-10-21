// Title: Copy only cell formatting from range O1:O5 to P1:P5 in a password‑protected Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy cell style from one range to another in a protected workbook | How to load an encrypted Excel file with a password and transfer formatting using Aspose.Cells | Copy only formatting between columns O and P in a password‑protected .xlsx with Aspose.Cells | C# Aspose.Cells example for copying range styles without moving data in a secured workbook | Load password‑protected workbook and apply source range formatting to destination range Aspose.Cells
// Tags: copy range formatting Aspose.Cells .NET | load password protected workbook Aspose.Cells | range style transfer Excel .NET | CopyStyle method Aspose.Cells | protected workbook LoadOptions C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads a password‑protected Excel workbook using LoadOptions, selects the source range O1:O5 and the destination range P1:P5 on the first worksheet, copies only the cell formatting with the CopyStyle method, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "ProtectedWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";
            const string password = "yourPassword"; // replace with actual password

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the password‑protected workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Define the source and destination ranges (O1:O5 -> P1:P5)
            Aspose.Cells.Range sourceRange = sheet.Cells.CreateRange("O1", "O5");
            Aspose.Cells.Range destRange = sheet.Cells.CreateRange("P1", "P5");

            // Copy only the formatting (styles) from source to destination
            destRange.CopyStyle(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
