// Title: Open a password‑protected XLSX with Aspose.Cells LightCells API in C#
// Description: Shows how to decrypt and load an encrypted XLSX workbook using Aspose.Cells LightCells API by assigning the Password property in LoadOptions, then walks through each worksheet and cell to print the address and value, with robust handling for missing files and invalid passwords.
// Keywords: Aspose.Cells LightCells password | C# load encrypted XLSX | LoadOptions Password property | read protected Excel cells .NET | open encrypted workbook Aspose | LightCells API example | password protected Excel Aspose.Cells | Excel decryption C#
// Common Searches: Aspose.Cells LightCells open password protected Excel | C# load encrypted XLSX with LoadOptions | How to read cells from a protected workbook using Aspose.Cells | LightCells API password example | Decrypt Excel file Aspose.Cells .NET | LoadOptions Password usage
// Developer Intent: Load an encrypted XLSX workbook via LightCells API and read its cell data in C#.
// Use Cases: Extract data from secured Excel reports for analytics pipelines. | Validate the contents of a password‑protected spreadsheet before further processing. | Migrate data from encrypted workbooks to a database or another file format. | Automate auditing of protected Excel files in enterprise workflows.
// AI Prompts: Generate C# code that opens a password‑protected XLSX file with Aspose.Cells LightCells API and prints every cell address and value. | Explain how to detect and handle an incorrect password exception when loading an encrypted workbook using LoadOptions. | Adapt the sample to stream the protected workbook from a MemoryStream instead of a file path while using LightCells API.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to decrypt and load an encrypted XLSX workbook using Aspose.Cells LightCells API by assigning the Password property in LoadOptions, then walks through each worksheet and cell to print the address and value, with robust handling for missing files and invalid passwords.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLSX file
        string filePath = "protected.xlsx";

        // Password used to encrypt the workbook
        string password = "test";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File \"{filePath}\" not found.");
            return;
        }

        try
        {
            // Set load options with the workbook password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Iterate through all worksheets and their cells
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                foreach (Cell cell in cells)
                {
                    // Output cell address (e.g., "A1") and its value
                    Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                }
            }
        }
        catch (Exception ex)
        {
            // Handle any errors that occur during loading or processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
