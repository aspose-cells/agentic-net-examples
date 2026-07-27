// Title: Open a password‑protected Excel workbook with Aspose.Cells LoadOptions (C#)
// Description: Demonstrates how to load an encrypted Excel file by setting LoadOptions.Password, verify the file's existence, read a cell value, and handle both Aspose.Cells‑specific and generic exceptions.
// Keywords: Aspose.Cells | C# | LoadOptions password | open encrypted Excel | password protected workbook | Workbook constructor | CellsException handling
// Common Searches: Aspose.Cells open encrypted workbook C# | LoadOptions Password example | Read cell from password protected Excel .NET | Handle invalid password Aspose.Cells | Check if Excel file exists before loading
// Developer Intent: Load a password‑protected Excel file and access its data using Aspose.Cells.
// Use Cases: Open an encrypted workbook and retrieve a specific cell value. | Validate a supplied password and report authentication failures. | Separate password errors from other Aspose.Cells exceptions.
// AI Prompts: Write C# code that opens a password‑protected Excel file with Aspose.Cells, updates cell B2, and saves the changes. | Show a try‑catch pattern for catching CellsException when the password is incorrect. | Create a console prompt that asks the user for a workbook password and loads the file using LoadOptions.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an encrypted Excel file by setting LoadOptions.Password, verify the file's existence, read a cell value, and handle both Aspose.Cells‑specific and generic exceptions.
class OpenEncryptedWorkbook
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Create LoadOptions and assign the password required to open the workbook
            LoadOptions loadOptions = new LoadOptions
            {
                Password = "myPassword"
            };

            // Load the workbook using the LoadOptions containing the password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example operation: read and display the value of cell A1 from the first worksheet
            Console.WriteLine("Cell A1 value: " + workbook.Worksheets[0].Cells["A1"].Value);
        }
        catch (CellsException ex)
        {
            // Handles errors related to Aspose.Cells operations (e.g., invalid password)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
