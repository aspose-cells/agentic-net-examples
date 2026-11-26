// Title: Add try‑catch to handle malformed JSON errors during Aspose.Cells JsonUtility.ImportData in C#
// Common Searches: how to catch exceptions from JsonUtility.ImportData when JSON is invalid in C# | Aspose.Cells JSON import malformed file error handling example | C# try catch for Aspose.Cells JsonUtility malformed JSON | handle JsonUtility.ImportData exception and continue processing in Aspose.Cells | detect and report invalid JSON while importing to Excel with Aspose.Cells
// Tags: Aspose.Cells JsonUtility.ImportData exception handling | C# malformed JSON import to Excel | Aspose.Cells JSON to worksheet error handling | try-catch JsonUtility.ImportData C# | import invalid JSON Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// The sample reads a JSON file, then uses Aspose.Cells JsonUtility.ImportData with JsonLayoutOptions to load the data into a worksheet. File reading, JSON import, and workbook saving are wrapped in separate try‑catch blocks that capture malformed‑JSON exceptions and other I/O errors, outputting clear error messages.
class Program
{
    static void Main(string[] args)
    {
        // Determine the JSON source file (first argument or default)
        string jsonFilePath = args.Length > 0 ? args[0] : "input.json";

        if (!File.Exists(jsonFilePath))
        {
            Console.WriteLine($"Error: JSON file not found at path \"{jsonFilePath}\".");
            return;
        }

        string jsonContent;
        try
        {
            jsonContent = File.ReadAllText(jsonFilePath);
        }
        catch (Exception readEx)
        {
            Console.WriteLine($"Error reading JSON file: {readEx.Message}");
            return;
        }

        // Prepare a workbook and target cells
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Configure import options (example settings)
        JsonLayoutOptions importOptions = new JsonLayoutOptions
        {
            ArrayAsTable = true,
            IgnoreNull = true
        };

        // Attempt to import the JSON data; catch malformed JSON exceptions
        try
        {
            // Import starts at row 0, column 0 (cell A1)
            JsonUtility.ImportData(jsonContent, cells, 0, 0, importOptions);
            Console.WriteLine("JSON data imported successfully.");
        }
        catch (Exception importEx)
        {
            // This block will capture any exception thrown due to malformed JSON
            Console.WriteLine($"Error importing JSON data: {importEx.Message}");
            // Optionally, you could log importEx.StackTrace for deeper diagnostics
            return;
        }

        // Save the resulting workbook to verify the import succeeded
        string outputPath = "output.xlsx";
        try
        {
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Error saving workbook: {saveEx.Message}");
        }
    }
}
