// Title: Encrypt an Excel workbook with a Unicode (accented & emoji) password and verify decryption using Aspose.Cells for .NET
// Common Searches: Aspose.Cells encrypt Excel workbook with Unicode password C# | How to open a password‑protected .xlsx that contains an emoji using Aspose.Cells | Set workbook password with accented characters in Aspose.Cells .NET | Load encrypted Excel file with special‑character password via LoadOptions C# | Confirm decryption of a workbook after saving it with a non‑ASCII password in Aspose.Cells
// Tags: non‑ASCII password workbook encryption Aspose.Cells | LoadOptions password for encrypted xlsx Aspose.Cells | verify workbook decryption Aspose.Cells .NET | protect Excel file using special‑character password .NET | emoji‑based password support Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a Workbook, writes data to cell A1, assigns a password containing accented letters and an emoji via wb.Settings.Password, saves the file as an encrypted .xlsx, then reloads it with a LoadOptions object that supplies the same Unicode password, and finally reads cell A1 to confirm successful decryption.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some data.
        Workbook wb = new Workbook();
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Unicode!");

        // Set a password containing Unicode characters.
        // Example password includes accented letters and an emoji.
        wb.Settings.Password = "Päßwörd🔐";

        // Save the encrypted workbook.
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);

        // Load the workbook using the same password.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = "Päßwörd🔐"
        };
        Workbook loadedWb = new Workbook(encryptedPath, loadOptions);

        // Verify that the workbook was decrypted successfully.
        string cellValue = loadedWb.Worksheets[0].Cells["A1"].StringValue;
        Console.WriteLine("Decrypted cell A1 value: " + cellValue);
    }
}
