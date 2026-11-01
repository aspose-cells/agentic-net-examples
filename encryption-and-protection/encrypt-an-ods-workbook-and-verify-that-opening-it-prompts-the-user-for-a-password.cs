// Title: Encrypt an ODS workbook with a password and verify the password prompt on opening using Aspose.Cells for .NET
// Common Searches: how to set a password for an ODS file with Aspose.Cells in C# | detect password required exception when opening encrypted ODS using Aspose.Cells | load password‑protected ODS workbook with LoadOptions in Aspose.Cells C# | example code for encrypting ODS and prompting for password on open with Aspose.Cells
// Tags: Aspose.Cells encrypt ODS workbook C# | Workbook.Settings.Password ODS encryption | LoadOptions.Password open protected ODS | CellsException handling password required ODS | verify password prompt ODS Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, assigns a password via Workbook.Settings.Password, saves it as an ODS file, attempts to open it without a password to capture the password‑required exception, and finally reloads the file using LoadOptions.Password to confirm successful decryption.
class OdsEncryptionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set a password for the ODS file
            workbook.Settings.Password = "Secret123";

            // Save the workbook as ODS with encryption
            string odsPath = "EncryptedSample.ods";
            workbook.Save(odsPath);

            // Attempt to load the encrypted ODS without providing a password
            try
            {
                if (File.Exists(odsPath))
                {
                    // This should throw because the file is password‑protected
                    Workbook loadWithoutPwd = new Workbook(odsPath);
                    Console.WriteLine("Loaded without password (unexpected).");
                }
                else
                {
                    Console.WriteLine($"File not found: {odsPath}");
                }
            }
            catch (CellsException ex)
            {
                // Older Aspose.Cells versions may not expose ErrorCode; inspect the message instead
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("Password required to open the workbook (as expected).");
                }
                else
                {
                    Console.WriteLine($"CellsException while loading without password: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error while loading without password: {ex.Message}");
            }

            // Load the encrypted ODS providing the correct password
            try
            {
                if (File.Exists(odsPath))
                {
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Ods)
                    {
                        Password = "Secret123"
                    };

                    Workbook loadWithPwd = new Workbook(odsPath, loadOptions);
                    Console.WriteLine("Workbook opened successfully with password.");
                }
                else
                {
                    Console.WriteLine($"File not found: {odsPath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook with password: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"General error: {ex.Message}");
        }
    }
}
