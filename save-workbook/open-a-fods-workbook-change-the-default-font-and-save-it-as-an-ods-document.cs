// Title: How to load a FODS workbook, change its default font to Arial, and save it as an ODS file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells change default workbook font before saving to ODS | Load FODS file and export to ODS with custom default style in .NET | Aspose.Cells example for setting default font of a spreadsheet and converting FODS to ODS | Programmatically convert FODS to ODS while applying Arial as default font using C# | Handle missing FODS file error when converting to ODS with Aspose.Cells
// Tags: load FODS workbook Aspose.Cells | modify default workbook font C# | export workbook to ODS format | FODS to ODS conversion Aspose.Cells | default style change Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input FODS file, loads it into an Aspose.Cells Workbook, updates the workbook's default style to use the Arial font, and then saves the workbook as an ODS document, with error handling for file access and conversion issues.
class Program
{
    static void Main()
    {
        const string inputFile = "input.fods";
        const string outputFile = "output.ods";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file '{inputFile}' was not found.");
            return;
        }

        try
        {
            // Load the FODS workbook
            Workbook workbook = new Workbook(inputFile);

            // Change the default font for the workbook
            workbook.DefaultStyle.Font.Name = "Arial";

            // Save the workbook as an ODS document
            workbook.Save(outputFile, SaveFormat.Ods);

            Console.WriteLine($"Workbook successfully saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving, Aspose.Cells specific issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
