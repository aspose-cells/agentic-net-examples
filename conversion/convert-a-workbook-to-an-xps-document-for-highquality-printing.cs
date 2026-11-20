// Title: Convert an Excel workbook to an XPS document for high‑quality printing with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to export an .xlsx file to XPS for printing | save Excel workbook as XPS using Aspose.Cells .NET API | generate high‑resolution XPS from Excel with Aspose.Cells C# example | Aspose.Cells LoadOptions password protected workbook then save as XPS
// Tags: Aspose.Cells XPS export | Excel to XPS conversion .NET | Workbook to XPS conversion | high‑resolution XPS printing from Excel | LoadOptions password protected workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input .xlsx file, loads it into an Aspose.Cells Workbook (optionally using LoadOptions for passwords), and saves the workbook as an XPS document for high‑quality printing, handling both Aspose.Cells‑specific and general exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xps";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the workbook; LoadOptions can be used to specify a password if needed
            LoadOptions loadOptions = new LoadOptions(); // No password provided
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook as an XPS document for high‑quality printing
            workbook.Save(outputPath, SaveFormat.Xps);
            Console.WriteLine($"Workbook successfully saved as XPS to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handles Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
