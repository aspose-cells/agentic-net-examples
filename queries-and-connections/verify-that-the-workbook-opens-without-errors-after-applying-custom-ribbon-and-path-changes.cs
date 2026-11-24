// Title: Validate that an Excel workbook opens after adding a custom document property and saving with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells verify that a saved workbook can be opened without errors | how to add a custom document property to an Excel file using Aspose.Cells and test loading | Aspose.Cells .NET ensure output workbook opens after modifying properties | validate workbook integrity after saving with Aspose.Cells in C# | C# code to load, modify, save, and re-open an Excel workbook using Aspose.Cells
// Tags: add custom document property Aspose.Cells C# | save workbook with custom property Aspose.Cells | verify workbook opens after save Aspose.Cells | create output directory before workbook save Aspose.Cells | handle missing CustomUIXml Aspose.Cells version issue

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, adds a custom document property named "CustomPath", ensures the output folder exists, saves the workbook as a new file, and then reloads the saved file to confirm it opens without exceptions, handling missing files and version‑specific API limitations.
class Program
{
    static void Main()
    {
        // Paths for input and output workbooks
        string originalPath = "input.xlsx";
        string modifiedPath = "output.xlsx";

        // Verify that the input file exists before attempting to load it
        if (!File.Exists(originalPath))
        {
            Console.WriteLine($"Input file not found: {originalPath}");
            return;
        }

        try
        {
            // Load the original workbook
            Workbook workbook = new Workbook(originalPath);

            // NOTE: The CustomUIXml property is not available in this version of Aspose.Cells.
            // If you need to add custom ribbon UI, use the appropriate API for your library version.

            // Store a custom document property (example: a custom path)
            workbook.CustomDocumentProperties.Add("CustomPath", @"C:\MyFolder\Resources");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(modifiedPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the customizations
            workbook.Save(modifiedPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
            return;
        }

        // Verify that the modified workbook opens without errors
        try
        {
            Workbook verifyWorkbook = new Workbook(modifiedPath);
            Console.WriteLine("Workbook opened successfully after applying customizations.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error opening modified workbook: {ex.Message}");
        }
    }
}
