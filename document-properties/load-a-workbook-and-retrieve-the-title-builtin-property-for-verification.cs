// Title: How to read the Title built‑in document property from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read Excel built‑in Title property from workbook | Retrieve document title from .xlsx using Aspose.Cells API in .NET | Get built‑in document properties of an Excel file with Aspose.Cells C# | Load password‑protected Excel file and access Title property using Aspose.Cells | Aspose.Cells example for reading workbook BuiltInDocumentProperties.Title in C#
// Tags: Aspose.Cells read built‑in document properties C# | load Excel workbook with Aspose.Cells .NET | extract Title metadata from .xlsx using Aspose.Cells | password protected workbook loading Aspose.Cells | access BuiltInDocumentProperties.Title C#

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the existence of an input.xlsx file, loads it into an Aspose.Cells Workbook (optionally supplying a password via LoadOptions), reads the Title built‑in document property through workbook.BuiltInDocumentProperties.Title, and prints the value to the console while handling CellsException and generic exceptions.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "input.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook (add password here if required)
            var loadOptions = new LoadOptions(); // Password can be set via loadOptions.Password = "yourPassword";
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Retrieve the Title built‑in property
            string title = workbook.BuiltInDocumentProperties.Title;

            // Output the title for verification
            Console.WriteLine("Title property: " + title);
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password required)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
