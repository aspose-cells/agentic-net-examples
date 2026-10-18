// Title: Detect if an Office Open XML (.xlsx) workbook is encrypted and report its status using Aspose.Cells for .NET
// Common Searches: how to check if an .xlsx file is password protected with Aspose.Cells in C# | C# code to detect encrypted Excel workbook without providing a password using Aspose.Cells | using Aspose.Cells LoadOptions to determine encryption status of an Excel file | catch encryption exception when loading a protected workbook with Aspose.Cells .NET | audit Excel file encryption by identifying encrypted workbooks via Aspose.Cells
// Tags: Aspose.Cells detect encrypted xlsx | LoadOptions password handling Aspose.Cells | C# check Excel workbook encryption status | exception based encryption detection Aspose.Cells | audit Excel file encryption Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of an .xlsx file, configures LoadOptions (optionally with a password), attempts to load the workbook using Aspose.Cells, and uses exception handling to determine whether the file is encrypted or the password is missing, then outputs the appropriate status message.
class EncryptionAudit
{
    static void Main()
    {
        // Path to the Office Open XML file (e.g., .xlsx)
        string filePath = "encrypted.xlsx";

        // Optional: provide the password if known; leave empty if not
        string password = "yourPassword";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Configure load options, including the password if supplied
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        if (!string.IsNullOrEmpty(password))
        {
            loadOptions.Password = password;
        }

        try
        {
            // Attempt to load the workbook with the provided options
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook loaded successfully. The file is either not encrypted or the correct password was provided.");
        }
        catch (Exception ex)
        {
            // Detect encryption‑related errors via the exception message
            if (ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 ||
                ex.Message.IndexOf("encrypted", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The file is encrypted, but a valid password was not provided.");
            }
            else
            {
                Console.WriteLine($"An error occurred while loading the workbook: {ex.Message}");
            }
        }
    }
}
