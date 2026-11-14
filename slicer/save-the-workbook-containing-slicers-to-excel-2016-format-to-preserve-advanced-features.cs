// Title: Save an Excel workbook that contains slicers to Excel 2016 (XLSX) format using Aspose.Cells for .NET
// Common Searches: asp.net save workbook with slicers to xlsx preserving slicer functionality | c# Aspose.Cells export Excel file with slicers to Excel 2016 format | using OoxmlSaveOptions to keep slicers when saving as XLSX in .NET | convert existing .xlsx containing slicers to Excel 2016 compatible file with Aspose.Cells | how to retain slicer objects after saving workbook with Aspose.Cells
// Tags: Aspose.Cells OoxmlSaveOptions XLSX | preserve slicers Aspose.Cells | save workbook with slicers C# | Excel 2016 compatibility Aspose.Cells | retain advanced features when exporting XLSX

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook that includes slicers, configures OoxmlSaveOptions with SaveFormat.Xlsx to maintain those slicers, and saves the file as an Excel 2016‑compatible XLSX, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "InputWithSlicers.xlsx";
            const string outputPath = "Output_Excel2016.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook that contains slicers
            Workbook workbook = new Workbook(inputPath);

            // Configure save options for XLSX format (preserves slicers)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook; slicers and other advanced features are retained
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
