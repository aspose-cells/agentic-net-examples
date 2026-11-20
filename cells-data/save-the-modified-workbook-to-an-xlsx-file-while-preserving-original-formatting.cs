// Title: C# – Save a modified Excel workbook to XLSX while preserving original formatting using Aspose.Cells
// Common Searches: C# Aspose.Cells save workbook to XLSX without losing cell styles | how to keep original Excel formatting when updating a workbook with Aspose.Cells .NET | Aspose.Cells .NET modify cell A1 and save preserving existing styles | load existing XLSX and save with same formatting using Aspose.Cells C# | save modified workbook while retaining original Excel formatting Aspose.Cells
// Tags: Aspose.Cells save workbook preserving original formatting | C# Aspose.Cells SaveFormat.Xlsx with style retention | load existing XLSX and save without formatting loss Aspose.Cells | modify cell and retain Excel styles using Aspose.Cells .NET | Aspose.Cells workbook Save method preserving formatting

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file (or creates a new workbook if the file is missing), updates cell A1, and then saves the workbook to a new XLSX file using Aspose.Cells' Save method with SaveFormat.Xlsx, ensuring that all original formatting and styles are retained.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            Workbook workbook = null;

            // Load existing workbook if it exists
            if (File.Exists(inputPath))
            {
                try
                {
                    // Attempt to open the workbook (no password)
                    workbook = new Workbook(inputPath);
                }
                catch (CellsException)
                {
                    // Likely a password‑protected file
                    Console.Error.WriteLine("The workbook is password protected. Provide a valid password to open it.");
                    return;
                }
            }
            else
            {
                // Create a new workbook with a default worksheet
                workbook = new Workbook();
            }

            // Example modification: change the value of cell A1 in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Modified");

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
