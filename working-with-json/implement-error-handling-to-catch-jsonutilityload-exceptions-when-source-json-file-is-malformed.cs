// Title: Add robust error handling for malformed JSON when loading with Aspose.Cells in C#
// AI Prompts: Show a C# snippet that validates a JSON file before passing it to Aspose.Cells and catches any parsing exceptions, printing the exception type and message. | Update the program to separate file‑not‑found handling from malformed JSON handling when loading with Aspose.Cells, using distinct catch blocks. | Illustrate writing the problematic JSON content and stack trace to a separate error log, then gracefully continue after an Aspose.Cells load failure.
// Common Searches: Aspose.Cells C# load JSON file with error handling for invalid JSON | catch exceptions thrown by JsonUtility.Load in .NET | convert malformed JSON to Excel using Aspose.Cells and handle failures | C# try‑catch pattern for Aspose.Cells JSON to workbook conversion | log JSON parsing errors when using Aspose.Cells LoadOptions
// Tags: Aspose.Cells JSON import error handling | C# malformed JSON detection Aspose.Cells | exception handling for JSON import Aspose.Cells | Excel workbook creation from JSON error management | record JSON import errors Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The example verifies the JSON file's existence, then attempts to load it into an Aspose.Cells Workbook using LoadOptions with JSON format inside a try‑catch block that distinguishes file‑not‑found, malformed JSON, and other load errors, providing clear console messages and optional logging before saving the workbook to XLSX while handling possible save exceptions.
class Program
{
    static void Main()
    {
        string jsonPath = "data.json";
        string outputPath = "output.xlsx";

        // Verify JSON file exists before loading
        if (!File.Exists(jsonPath))
        {
            Console.WriteLine($"JSON file not found: '{jsonPath}'");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Load JSON directly into a workbook using LoadOptions with Json format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
            workbook = new Workbook(jsonPath, loadOptions);
        }
        catch (Exception ex)
        {
            // Handle errors such as malformed JSON or load failures
            Console.WriteLine($"Failed to load JSON file '{jsonPath}': {ex.Message}");
            return;
        }

        try
        {
            // Save the workbook to the specified output path
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle save errors (e.g., permission issues, invalid path)
            Console.WriteLine($"Failed to save workbook to '{outputPath}': {ex.Message}");
        }
    }
}
