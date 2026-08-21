// Title: Open a Password‑Protected Excel Workbook with Aspose.Cells LoadOptions in C#
// Description: Shows how to load an encrypted .xlsx file in C# by assigning LoadOptions.Password, verify the file’s presence, read a cell value, check the workbook’s encryption flag, and handle password‑related exceptions with Aspose.Cells.
// Keywords: Aspose.Cells | C# load encrypted workbook | LoadOptions.Password | password protected Excel file | read encrypted .xlsx | CellsException handling | Aspose.Cells .NET example | open encrypted Excel workbook | Excel password C# | encrypted workbook code sample
// Common Searches: Aspose.Cells open password protected Excel C# | LoadOptions.Password usage example | Read cell from encrypted workbook using Aspose.Cells | Check if workbook is encrypted Aspose.Cells | Handle invalid password exception Aspose.Cells
// Developer Intent: Load a password‑protected Excel file and access its contents with Aspose.Cells.
// Use Cases: Open an encrypted .xlsx, read a specific cell, and confirm successful decryption. | Detect and report an invalid password by catching CellsException. | Integrate secure workbook loading into automated data‑processing pipelines.
// AI Prompts: Generate C# code that opens a password‑protected Excel workbook using Aspose.Cells LoadOptions and reads a cell value. | Explain how to catch CellsException for an incorrect password when loading an encrypted workbook with Aspose.Cells. | Show how to determine whether a loaded workbook is encrypted and log the result in C#.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an encrypted .xlsx file in C# by assigning LoadOptions.Password, verify the file’s presence, read a cell value, check the workbook’s encryption flag, and handle password‑related exceptions with Aspose.Cells.
class OpenEncryptedWorkbook
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File \"{filePath}\" not found.");
            return;
        }

        try
        {
            // Create LoadOptions and set the password required to open the workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "myPassword"
            };

            // Load the password‑protected workbook using the LoadOptions
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Access a cell to verify that the workbook was opened successfully
            Console.WriteLine("Cell C6 value: " + workbook.Worksheets[0].Cells["C6"].Value);

            // Optional: display whether the workbook is encrypted
            Console.WriteLine("IsEncrypted: " + workbook.Settings.IsEncrypted);
        }
        catch (CellsException ex)
        {
            // Handle errors related to loading the workbook (e.g., invalid password)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
