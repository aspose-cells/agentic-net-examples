// Title: Load a password‑protected XLSX workbook with Aspose.Cells LightCells API (C#)
// Description: Demonstrates how to open an encrypted Excel file in .NET by creating a LoadOptions object, setting its Password property, and passing it to the Workbook constructor. The sample checks file existence, prints the first worksheet name, and handles errors such as wrong passwords or corrupted files.
// Keywords: Aspose.Cells load password protected XLSX | LoadOptions.Password C# | open encrypted Excel file .NET | LightCells API password | read protected workbook Aspose
// Common Searches: Aspose.Cells open password protected Excel | C# load encrypted XLSX with LoadOptions | How to read a protected workbook using Aspose.Cells | LightCells API password example
// Developer Intent: Open a password‑protected XLSX file by supplying the decryption password through LoadOptions.
// Use Cases: Access a secured workbook after verifying the file path. | Retrieve worksheet names or data from an encrypted Excel file. | Gracefully handle invalid passwords or corrupted files with exception handling.
// AI Prompts: Show C# code that loads a password‑protected XLSX using Aspose.Cells LightCells API and LoadOptions.Password. | Explain how to configure LoadOptions for decryption and read the first sheet name from a protected workbook. | Provide an example that iterates all worksheets in an encrypted Excel file and catches wrong‑password errors.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to open an encrypted Excel file in .NET by creating a LoadOptions object, setting its Password property, and passing it to the Workbook constructor. The sample checks file existence, prints the first worksheet name, and handles errors such as wrong passwords or corrupted files.
class LoadPasswordProtectedWorkbook
{
    static void Main()
    {
        // Path to the password‑protected XLSX file
        string filePath = "protected.xlsx";

        // The password used to encrypt the workbook
        string password = "test";

        try
        {
            // Verify that the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File \"{filePath}\" not found.");
                return;
            }

            // Create LoadOptions and assign the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the workbook using the standard API with the specified LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example verification: output the name of the first worksheet
            Console.WriteLine("Workbook loaded successfully. First sheet name: " + workbook.Worksheets[0].Name);
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions (e.g., incorrect password, corrupted file)
            Console.WriteLine("An error occurred while loading the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
