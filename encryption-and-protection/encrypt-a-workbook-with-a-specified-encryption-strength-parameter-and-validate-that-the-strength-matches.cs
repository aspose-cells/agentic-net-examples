// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and confirm successful decryption in C#
// Common Searches: how to apply password protection to an XLSX file with Aspose.Cells in C# | load a password‑protected workbook using Aspose.Cells LoadOptions C# | verify that encrypted Excel data can be read after decryption with Aspose.Cells | save an encrypted workbook and check cell value after opening in .NET | Aspose.Cells encrypt workbook and validate password protection programmatically
// Tags: password protection workbook Aspose.Cells .NET | save encrypted XLSX Aspose.Cells | load password protected workbook LoadOptions Aspose.Cells | validate decrypted cell value Aspose.Cells C# | Aspose.Cells workbook encryption example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, writes sample data, sets a password via workbook.Settings.Password, saves the file as an encrypted XLSX, then reloads it with LoadOptions using the same password and confirms that the original cell value is correctly retrieved.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Define password for encryption
            const string password = "SecurePassword123";

            // Apply password protection (encryption) to the workbook
            workbook.Settings.Password = password;

            // Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved and encrypted at: {Path.GetFullPath(filePath)}");

            // Verify the file exists before attempting to load
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Encrypted workbook file not found.", filePath);

            // Load the workbook using the password to verify encryption
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Simple validation: ensure we can read a cell value after decryption
            string cellValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            if (cellValue != "Sample Data")
                throw new Exception("Decrypted data does not match expected value.");

            Console.WriteLine("Encryption validated successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
