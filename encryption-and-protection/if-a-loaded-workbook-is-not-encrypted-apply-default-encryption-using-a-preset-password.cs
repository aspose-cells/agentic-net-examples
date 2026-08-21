// Title: C# – Apply default password encryption to an unencrypted Excel workbook with Aspose.Cells
// Description: Loads a workbook, checks if it is already encrypted, and if not, sets a preset password, enables default encryption, optionally applies a strong cryptographic provider (128‑bit), then saves the protected file.
// Keywords: Aspose.Cells encrypt workbook C# | Excel password protection .NET | Workbook.Settings.IsEncrypted | SetEncryptionOptions strong encryption | default encryption Aspose.Cells
// Common Searches: encrypt existing Excel file Aspose.Cells C# | check workbook encryption before applying password | set default encryption with preset password Aspose.Cells | strong cryptographic provider for Excel files .NET
// Developer Intent: Add password protection only when a workbook lacks encryption, using Aspose.Cells default encryption settings.
// Use Cases: Secure user‑uploaded spreadsheets on a web server without double‑encrypting already protected files. | Generate automated reports that are automatically password‑protected if the source workbook is plain. | Enforce organizational data‑security policies by applying 128‑bit strong encryption to Excel outputs.
// AI Prompts: Write C# code that opens an Excel file with Aspose.Cells, verifies Workbook.Settings.IsEncrypted, and applies a preset password with default encryption and a 128‑bit strong cryptographic provider. | Create a reusable method (inputPath, outputPath, password) that encrypts a workbook only when it is not already encrypted and returns true if encryption was performed. | Explain how to configure different encryption algorithms and key lengths using Aspose.Cells SetEncryptionOptions.

using Aspose.Cells;

// Loads a workbook, checks if it is already encrypted, and if not, sets a preset password, enables default encryption, optionally applies a strong cryptographic provider (128‑bit), then saves the protected file.
class Program
{
    static void Main()
    {
        // Paths and preset password
        string inputPath = "input.xlsx";
        string outputPath = "output_encrypted.xlsx";
        string presetPassword = "MySecretPassword";

        // Load the workbook (existing lifecycle rule)
        Workbook workbook = new Workbook(inputPath);

        // If the workbook is not encrypted, apply default encryption
        if (!workbook.Settings.IsEncrypted)
        {
            // Set the password that will protect the file
            workbook.Settings.Password = presetPassword;

            // Mark that default encryption should be used
            workbook.Settings.IsDefaultEncrypted = true;

            // Optional: specify a strong encryption algorithm and key length
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);
        }

        // Save the workbook (existing lifecycle rule)
        workbook.Save(outputPath);
    }
}
