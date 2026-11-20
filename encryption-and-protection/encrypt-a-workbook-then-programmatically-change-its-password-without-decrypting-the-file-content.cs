// Title: Change the password of an encrypted Excel workbook programmatically with Aspose.Cells for .NET
// Common Searches: asp.net change password of encrypted .xlsx using Aspose.Cells | replace workbook encryption password without opening file in C# | load encrypted workbook with password and save with new password Aspose.Cells | update Excel file protection password programmatically Aspose.Cells .NET | change workbook Settings.Password after loading encrypted file Aspose.Cells
// Tags: Aspose.Cells workbook encryption password modification | Load encrypted .xlsx with LoadOptions.Password | Workbook.Settings.Password usage in C# | Programmatic Excel file password replacement | C# update encrypted workbook without decryption

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, encrypts it by setting Settings.Password, saves it, then reloads the encrypted file using LoadOptions.Password, assigns a new password to Settings.Password, and saves again, allowing the workbook's encryption password to be changed without decrypting its contents.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(42);

            // Encrypt the workbook with an initial password
            workbook.Settings.Password = "oldPassword";

            // Save the encrypted workbook
            string filePath = "EncryptedWorkbook.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // Ensure the file exists before attempting to load
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found.");

            // Load the encrypted workbook using the current password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "oldPassword"
            };
            Workbook loadedWorkbook = new Workbook(filePath, loadOptions);

            // Change the password by setting a new password on the workbook settings
            loadedWorkbook.Settings.Password = "newPassword";

            // Save the workbook with the new password (overwrites the original file)
            loadedWorkbook.Save(filePath, SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
