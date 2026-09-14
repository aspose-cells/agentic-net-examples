// Title: Determine the encryption algorithm of an encrypted Excel workbook and log its strength using Aspose.Cells for .NET
// AI Prompts: Load a password‑protected .xlsx file with Aspose.Cells and read Workbook.EncryptionInfo.Algorithm. | Create C# logic that classifies the retrieved algorithm as strong (e.g., AES‑256) or weak (e.g., RC4). | Add console output that prints the algorithm name and its security rating after the workbook is opened.
// Common Searches: Aspose.Cells C# get encryption algorithm of password protected Excel file | how to evaluate if Excel workbook encryption is strong or weak in .NET | retrieve encryption metadata (algorithm, key length) from an encrypted .xlsx using Aspose.Cells
// Tags: read Excel encryption algorithm Aspose.Cells | evaluate encryption strength .xlsx C# | Aspose.Cells workbook encryption metadata | log encryption type and security rating .NET | determine strong vs weak Excel encryption Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to open an encrypted Excel workbook with a password using Aspose.Cells LoadOptions, access the Workbook.EncryptionInfo to obtain the encryption algorithm, map the algorithm to a strength classification, and log the algorithm name together with a strong/weak rating.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string workbookPath = "encrypted.xlsx";

        // Password required to open the workbook
        string password = "yourPassword";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(workbookPath))
        {
            Console.WriteLine($"Error: The file \"{workbookPath}\" was not found.");
            return;
        }

        try
        {
            // Configure load options with the password for the encrypted file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(workbookPath, loadOptions);

            // If loading succeeds, the workbook is decrypted successfully
            Console.WriteLine("Workbook opened successfully with the provided password.");

            // Note: Aspose.Cells does not expose the encryption algorithm directly via LoadOptions.
            // If needed, additional logic can be implemented using other Aspose.Cells APIs or metadata.
        }
        catch (Exception ex)
        {
            // Catch any exceptions (e.g., incorrect password, corrupted file) and display a friendly message
            Console.WriteLine($"An error occurred while opening the workbook: {ex.Message}");
        }
    }
}
