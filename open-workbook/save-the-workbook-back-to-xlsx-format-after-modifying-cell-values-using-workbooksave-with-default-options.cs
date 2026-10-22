// Title: Save modified Excel workbook to XLSX with Aspose.Cells Workbook.Save (default options) in C#
// Common Searches: C# Aspose.Cells how to save a workbook after changing cell values | Saving an Excel file to a new path using Aspose.Cells Workbook.Save default settings | Aspose.Cells example for loading an existing XLSX, editing cells, and saving | How to handle exceptions when saving a workbook with Aspose.Cells in C# | Write updated workbook to output.xlsx using Aspose.Cells without specifying SaveOptions
// Tags: Aspose.Cells Workbook.Save to XLSX | C# edit and save Excel workbook | load Excel file and save with Aspose.Cells | exception handling for workbook save Aspose.Cells | default save options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing input.xlsx (or creates a new workbook if missing or protected), updates cells A1 and B2, and then saves the workbook to output.xlsx using Workbook.Save with default options, including basic error handling for both loading and saving operations.
class Program
{
    static void Main()
    {
        Workbook workbook;

        // Ensure the input file exists before attempting to load it
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        if (File.Exists(inputPath))
        {
            try
            {
                // Attempt to load the existing workbook
                workbook = new Workbook(inputPath);
            }
            catch (CellsException ex)
            {
                // Handle cases such as password‑protected files
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                Console.WriteLine("A new workbook will be created instead.");
                workbook = new Workbook();
            }
        }
        else
        {
            Console.WriteLine("Input file not found. A new workbook will be created.");
            workbook = new Workbook();
        }

        // Access the first worksheet (creates one if none exist)
        Worksheet sheet = workbook.Worksheets[0];

        // Modify cell values as needed
        sheet.Cells["A1"].PutValue("Sample Text");
        sheet.Cells["B2"].PutValue(12345);

        try
        {
            // Save the workbook to the specified output file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
