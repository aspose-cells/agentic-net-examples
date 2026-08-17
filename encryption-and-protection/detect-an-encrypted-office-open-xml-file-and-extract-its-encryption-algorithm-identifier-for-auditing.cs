// Title: Detect encrypted OOXML Excel file and identify its encryption algorithm using Aspose.Cells for .NET
// Description: This C# example shows how to verify a workbook's existence, detect whether an .xlsx file is encrypted with Aspose.Cells' FileFormatUtil, prompt for a password, and attempt to open the file via LoadOptions. It reports the encryption status and highlights that Aspose.Cells does not expose the encryption algorithm identifier directly, guiding auditors toward alternative strategies.
// Keywords: Aspose.Cells | C# | .xlsx encryption detection | OOXML encryption algorithm | FileFormatUtil IsEncrypted | LoadOptions password | audit Excel encryption | encryption algorithm identifier | detect encrypted workbook | Excel security audit
// Common Searches: How to check if an .xlsx file is encrypted with Aspose.Cells | Retrieve encryption algorithm of an encrypted Excel workbook in C# | Detect encrypted Office Open XML file using Aspose.Cells | Audit Excel file encryption with .NET | Aspose.Cells encryption algorithm identifier
// Developer Intent: Determine whether an Excel .xlsx file is encrypted and obtain its encryption algorithm identifier for compliance auditing.
// Use Cases: Validate incoming Excel documents for encryption before processing in an automated workflow. | Prompt users for a password, attempt to open an encrypted workbook, and handle invalid passwords gracefully. | Document the limitation that Aspose.Cells does not expose the encryption algorithm identifier, and suggest alternative auditing approaches.
// AI Prompts: Write C# code with Aspose.Cells that detects if an .xlsx file is encrypted and returns the encryption algorithm identifier for audit purposes. | Create a method that loads an encrypted OOXML workbook using a supplied password and extracts the encryption algorithm name, handling cases where the API lacks direct support. | Suggest a strategy to audit the encryption algorithm of an encrypted Excel file when Aspose.Cells cannot retrieve it directly.

using System;
using System.IO;
using Aspose.Cells;

// This C# example shows how to verify a workbook's existence, detect whether an .xlsx file is encrypted with Aspose.Cells' FileFormatUtil, prompt for a password, and attempt to open the file via LoadOptions. It reports the encryption status and highlights that Aspose.Cells does not expose the encryption algorithm identifier directly, guiding auditors toward alternative strategies.
class EncryptionAudit
{
    static void Main()
    {
        string filePath = "encrypted.xlsx";

        // Verify that the file exists before proceeding
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Detect file format and encryption status
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

            if (!formatInfo.IsEncrypted)
            {
                Console.WriteLine("The file is not encrypted.");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error detecting file format: {ex.Message}");
            return;
        }

        // Prompt user for the password
        Console.Write("Enter password to open the workbook: ");
        string password = Console.ReadLine() ?? string.Empty;

        try
        {
            // Load workbook with the supplied password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
            {
                Password = password
            };
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Aspose.Cells does not expose the encryption algorithm directly after loading.
            // Indicate that the workbook was opened successfully.
            Console.WriteLine("Workbook opened successfully.");
        }
        catch (Exception ex)
        {
            // Handle invalid password or other loading errors
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Invalid password provided.");
            }
            else
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
