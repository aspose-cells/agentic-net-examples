// Title: C# – Load a macro‑enabled XLSM workbook with Aspose.Cells and read the first worksheet
// Description: The example checks whether a specified .xlsm file exists, creates a minimal placeholder workbook if needed, loads the macro‑enabled workbook using Aspose.Cells, retrieves the first worksheet, prints its name, and handles any exceptions.
// Keywords: Aspose.Cells | C# | load XLSM | macro-enabled workbook | open Excel file | read first worksheet | create placeholder workbook | SaveFormat.Xlsm | exception handling | file path
// Common Searches: How to open an .xlsm file using Aspose.Cells C# | Aspose.Cells create placeholder XLSM if file missing | Get first worksheet name after loading workbook Aspose.Cells | Load macro-enabled workbook from path Aspose.Cells | C# example for reading XLSM with Aspose
// Developer Intent: Load an existing macro‑enabled XLSM file from a given path and retrieve the name of its first worksheet, optionally creating a placeholder file when the target does not exist.
// Use Cases: Open a macro‑enabled workbook to extract data or metadata. | Automatically generate a minimal .xlsm file when the expected file is absent. | Integrate workbook loading into automated scripts with robust error handling. | Display or log the name of the first worksheet for verification.
// AI Prompts: Generate C# code using Aspose.Cells to open an .xlsm file from a path, create a minimal placeholder workbook if the file does not exist, and print the first worksheet name with proper exception handling. | Show an example that checks for a macro‑enabled Excel file, creates a placeholder using SaveFormat.Xlsm when missing, loads it with Aspose.Cells, and outputs the name of the first sheet.

using System;
using System.IO;
using Aspose.Cells;

// The example checks whether a specified .xlsm file exists, creates a minimal placeholder workbook if needed, loads the macro‑enabled workbook using Aspose.Cells, retrieves the first worksheet, prints its name, and handles any exceptions.
class Program
{
    static void Main()
    {
        // Path to the macro-enabled workbook
        string filePath = @"C:\Data\MyMacroEnabledWorkbook.xlsm";

        try
        {
            // Ensure the file exists; create a minimal workbook if it does not
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                Workbook placeholder = new Workbook();
                placeholder.Worksheets[0].Name = "Sheet1";
                placeholder.Save(filePath, SaveFormat.Xlsm);
                Console.WriteLine($"Created placeholder workbook at: {filePath}");
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"Loaded workbook. First worksheet name: {firstSheet.Name}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
