// Title: Convert a CSV file to PDF with Aspose.Cells for .NET while handling missing EnableCellControls for Office Add‑In interactivity
// AI Prompts: Generate C# code that reads a CSV file using Aspose.Cells LoadOptions, validates the file's existence, and saves it as a PDF with comprehensive exception handling. | Explain why the EnableCellControls property may be unavailable in certain Aspose.Cells releases and suggest workarounds for preserving Office Add‑In interactive controls during PDF export. | Show how to enhance the CSV‑to‑PDF routine to log detailed error information to a file and return an exit status instead of writing messages to the console.
// Common Searches: Aspose.Cells C# convert CSV to PDF with error handling | How to preserve Office Add‑In controls when exporting a workbook to PDF using Aspose.Cells | EnableCellControls property missing in Aspose.Cells version alternatives | LoadOptions CSV example for Aspose.Cells .NET | Check file existence before converting CSV to PDF with Aspose.Cells
// Tags: Aspose.Cells CSV to PDF conversion C# | LoadOptions CSV Aspose.Cells | SaveFormat.Pdf Aspose.Cells usage | EnableCellControls unavailable Aspose.Cells | Office Add‑In interactive controls PDF export

using Aspose.Cells;
using System;
using System.IO;

// The example validates the presence of an input CSV file, loads it into an Aspose.Cells Workbook using LoadOptions with LoadFormat.Csv, and saves the workbook as a PDF. It includes try‑catch error handling and notes that the EnableCellControls property for preserving Office Add‑In interactive controls is not available in the referenced Aspose.Cells version.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.csv";
            const string outputPath = "output.pdf";

            // Verify that the input CSV file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the CSV file into a workbook using appropriate load options
            var loadOptions = new LoadOptions(LoadFormat.Csv);
            var workbook = new Workbook(inputPath, loadOptions);

            // NOTE: The property to preserve interactive controls (EnableCellControls) is not
            // available in the referenced Aspose.Cells version, so it is omitted.

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Successfully converted '{inputPath}' to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
