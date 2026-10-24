// Title: How to confirm that an Excel workbook is password‑protected after saving it with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check if a saved workbook is password protected | How to read Settings.IsProtected after encrypting an XLSX with Aspose.Cells | Load encrypted Excel file with password using Aspose.Cells LoadOptions and verify protection | Determine workbook encryption flag in .NET after calling Workbook.Save | Validate password protection flag of an Excel file created with Aspose.Cells
// Tags: Aspose.Cells verify workbook password protection | C# load encrypted XLSX with LoadOptions password | Settings.IsProtected check Aspose.Cells | Excel file encryption status check .NET | Workbook.Settings.Password usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a Workbook, sets a password via Settings.Password, saves it as an encrypted XLSX, reloads the file with LoadOptions using the same password, and reads Settings.IsProtected to confirm that the password‑protection flag is set.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set a password to protect the workbook
            workbook.Settings.Password = "SecretPassword";

            // Save the workbook as an encrypted file
            string filePath = "ProtectedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before attempting to load it
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "SecretPassword"
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Verify that the workbook is password protected
            bool isPasswordProtected = loadedWorkbook.Settings.IsProtected;
            Console.WriteLine("Password protection flag is set: " + isPasswordProtected);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
