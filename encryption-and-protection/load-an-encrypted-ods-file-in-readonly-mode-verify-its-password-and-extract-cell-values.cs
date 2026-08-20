// Title: Open an Encrypted ODS Workbook in Read‑Only Mode, Verify Its Password, and Extract Cell Values with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to confirm a password using FileFormatUtil.VerifyPassword, load a password‑protected ODS file with OdsLoadOptions in read‑only mode, and iterate over the populated range of the first worksheet to output non‑null cell values. Includes basic error handling for missing files and invalid passwords.
// Keywords: Aspose.Cells | C# | .NET | ODS | encrypted ODS | password protected ODS | verify ODS password | OdsLoadOptions | read‑only workbook | extract cell values | FileFormatUtil VerifyPassword
// Common Searches: C# open password protected ODS file Aspose.Cells | verify ODS workbook password before loading | read data from encrypted ODS spreadsheet .NET | Aspose.Cells OdsLoadOptions password example | how to check ODS file password C#
// Developer Intent: The developer needs to open a password‑protected ODS file without modifying it, ensure the supplied password is correct, and read the worksheet’s data.
// Use Cases: Validate a user‑entered password before processing a secured ODS report. | Extract values from a protected ODS template for calculations or reporting while keeping the source file read‑only. | Log all populated cells from an encrypted ODS workbook for audit or debugging purposes.
// AI Prompts: Generate C# code that uses Aspose.Cells to verify a password, load an encrypted ODS file in read‑only mode, and print each non‑null cell value. | Explain best practices for handling incorrect passwords and missing files when opening a password‑protected ODS workbook with Aspose.Cells. | Show how to retrieve the maximum data row and column of a worksheet after loading an ODS file with OdsLoadOptions.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to confirm a password using FileFormatUtil.VerifyPassword, load a password‑protected ODS file with OdsLoadOptions in read‑only mode, and iterate over the populated range of the first worksheet to output non‑null cell values. Includes basic error handling for missing files and invalid passwords.
class LoadEncryptedOds
{
    static void Main()
    {
        // Path to the encrypted ODS file and its password
        string filePath = "encrypted_file.ods";
        string password = "myPassword";

        // Verify that the file exists before proceeding
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Verify the password before attempting to load the workbook
            bool isPasswordValid;
            using (FileStream stream = File.OpenRead(filePath))
            {
                isPasswordValid = FileFormatUtil.VerifyPassword(stream, password);
            }
            Console.WriteLine($"Password verification result: {isPasswordValid}");

            if (!isPasswordValid)
            {
                Console.WriteLine("Invalid password. Unable to open the workbook.");
                return;
            }

            // Load the ODS file with the provided password using OdsLoadOptions
            OdsLoadOptions loadOptions = new OdsLoadOptions
            {
                Password = password // set password for loading
            };

            Workbook workbook = new Workbook(filePath, loadOptions); // read‑only load

            // Extract and display cell values from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var value = cells[row, col].Value;
                    if (value != null)
                    {
                        Console.WriteLine($"Cell {cells[row, col].Name}: {value}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
