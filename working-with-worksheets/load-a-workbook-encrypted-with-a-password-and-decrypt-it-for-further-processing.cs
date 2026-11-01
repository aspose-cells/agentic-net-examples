// Title: Decrypt a password‑protected Excel workbook and save it unencrypted using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells LoadOptions with the Password property to open an encrypted .xlsx file in C# and then save it without a password. | Write C# code that loads a password‑protected workbook, removes its protection in memory, and creates a plain copy. | Demonstrate how to open an encrypted Excel file with Aspose.Cells, process it, and export a decrypted version.
// Common Searches: how to open a password protected xlsx with aspose.cells c# | remove workbook password using Aspose.Cells LoadOptions | save decrypted copy of encrypted Excel file in .NET | Aspose.Cells decrypt Excel file without user prompt
// Tags: Aspose.Cells LoadOptions password decryption | C# open encrypted XLSX workbook | Workbook.Save without password Aspose.Cells | remove Excel file protection programmatically | process password‑protected workbook .NET

using System;
using Aspose.Cells;

// // Loads an encrypted Excel workbook using LoadOptions.Password, decrypts it in memory, and saves a new unprotected .xlsx file with Aspose.Cells.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string inputPath = "encrypted.xlsx";

        // Password used to encrypt the workbook
        string password = "myPassword";

        // Load the workbook with the password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // At this point the workbook is decrypted in memory and can be processed further
        // Example processing: (add your logic here)

        // Save the workbook without a password (decrypted file)
        string outputPath = "decrypted.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}
