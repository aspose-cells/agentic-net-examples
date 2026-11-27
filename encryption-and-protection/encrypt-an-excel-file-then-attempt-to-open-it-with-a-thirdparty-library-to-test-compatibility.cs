// Title: Encrypt an XLSX workbook with a password using Aspose.Cells for .NET and verify opening with correct and incorrect passwords
// Common Searches: how to set password protection on an Excel workbook using Aspose.Cells C# | load password protected xlsx file with Aspose.Cells LoadOptions example | catch exception when opening encrypted Excel file with wrong password Aspose.Cells | save encrypted XLSX to temporary folder using Aspose.Cells .NET | test compatibility of Aspose.Cells encrypted workbook with third‑party libraries
// Tags: Aspose.Cells workbook password encryption | C# save encrypted XLSX Aspose.Cells | LoadOptions password protected Excel opening | exception handling wrong password Aspose.Cells | temporary file encryption Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a temporary XLSX workbook, adds data, applies password protection via Workbook.Settings.Password, saves it, then attempts to open the file twice using LoadOptions—once with the correct password to read a cell value and once with an incorrect password to demonstrate proper exception handling.
class Program
{
    static void Main()
    {
        // Path for the temporary Excel file
        string filePath = Path.Combine(Path.GetTempPath(), "EncryptedSample.xlsx");
        string password = "Secret123";

        // -------------------------------------------------
        // Create a workbook with Aspose.Cells and encrypt it
        // -------------------------------------------------
        // Initialize a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");

        // Set password protection (encryption)
        workbook.Settings.Password = password; // Encrypt the file with the given password

        // Save the encrypted workbook
        workbook.Save(filePath, SaveFormat.Xlsx);
        Console.WriteLine($"Aspose.Cells: Encrypted file saved to {filePath}");

        // -------------------------------------------------
        // Attempt to open the encrypted file with Aspose.Cells
        // -------------------------------------------------
        if (!File.Exists(filePath))
        {
            Console.WriteLine("Error: Encrypted file not found.");
            return;
        }

        // 1. Try opening with the correct password (should succeed)
        try
        {
            LoadOptions correctLoadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = password
            };
            Workbook loadedWorkbook = new Workbook(filePath, correctLoadOptions);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Console.WriteLine("Aspose.Cells: Opened with correct password.");
            Console.WriteLine($"First cell value (A2): {loadedSheet.Cells["A2"].StringValue}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Aspose.Cells: Failed to open with correct password. Exception: {ex.Message}");
        }

        // 2. Try opening with an incorrect password (should fail)
        try
        {
            LoadOptions wrongLoadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "WrongPassword"
            };
            Workbook wrongWorkbook = new Workbook(filePath, wrongLoadOptions);
            // If we get here, the file opened unexpectedly
            Console.WriteLine("Aspose.Cells: Unexpectedly opened with wrong password.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Aspose.Cells: Correctly failed to open with wrong password. Exception: {ex.Message}");
        }

        // Cleanup (optional)
        // File.Delete(filePath);
    }
}
