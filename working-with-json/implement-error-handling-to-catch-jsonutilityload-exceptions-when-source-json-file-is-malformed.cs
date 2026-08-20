// Title: C# – Catch malformed JSON errors when loading a workbook with Aspose.Cells JsonLoadOptions
// Description: Demonstrates how to load a JSON file into an Aspose.Cells Workbook using JsonLoadOptions, then save it as Excel while handling malformed‑JSON scenarios. The example catches CellsException for FileCorrupted and InvalidData and provides a generic fallback for other errors.
// Keywords: Aspose.Cells JsonLoadOptions | C# JSON to Excel error handling | CellsException FileCorrupted | InvalidData exception Aspose | malformed JSON Aspose.Cells | load JSON workbook C# | exception handling Aspose.Cells | JSON parsing error C#
// Common Searches: how to catch malformed JSON with Aspose.Cells C# | Aspose.Cells JsonLoadOptions exception types | C# example catching CellsException for JSON load | error handling when converting JSON to Excel using Aspose | Aspose.Cells JSON file corrupted error
// Developer Intent: Add robust try‑catch logic to capture JsonUtility.Load failures caused by malformed JSON when creating a Workbook with Aspose.Cells.
// Use Cases: Validate a JSON source before conversion and show a clear message if the file is corrupted. | Log detailed CellsException information to aid troubleshooting of data import problems. | Switch to an alternative data source automatically when the primary JSON cannot be parsed.
// AI Prompts: Write a reusable C# method that loads JSON into an Aspose.Cells Workbook with JsonLoadOptions and returns detailed error information for malformed files. | Generate code to log CellsException details (code, message, stack trace) to a file when JSON loading fails. | Explain the difference between CellsException codes FileCorrupted and InvalidData in the context of JSON import.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells; // for ExceptionType

// Demonstrates how to load a JSON file into an Aspose.Cells Workbook using JsonLoadOptions, then save it as Excel while handling malformed‑JSON scenarios. The example catches CellsException for FileCorrupted and InvalidData and provides a generic fallback for other errors.
class JsonLoadWithErrorHandling
{
    static void Main()
    {
        // Path to the source JSON file
        string jsonPath = "data.json";

        // Create JSON load options (customize as needed)
        JsonLoadOptions loadOptions = new JsonLoadOptions
        {
            KeptSchema = true
        };

        try
        {
            // Attempt to load the JSON file into a workbook
            Workbook workbook = new Workbook(jsonPath, loadOptions);

            // If loading succeeds, save the workbook to an Excel file
            workbook.Save("output.xlsx");
            Console.WriteLine("JSON loaded and saved successfully.");
        }
        // Catch Aspose.Cells specific exceptions that indicate malformed JSON
        catch (CellsException ex) when (ex.Code == ExceptionType.FileCorrupted ||
                                        ex.Code == ExceptionType.InvalidData)
        {
            Console.WriteLine($"Malformed JSON file: {ex.Message}");
        }
        // Fallback for any other unexpected exceptions
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading JSON: {ex.Message}");
        }
    }
}
