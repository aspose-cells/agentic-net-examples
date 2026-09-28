// Title: Encrypt an XLSX workbook with a 256‑bit password using Aspose.Cells and the Microsoft Strong Cryptographic Provider (C#)
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, sets a 256‑bit password via the Microsoft Strong Cryptographic Provider, and saves the file as an encrypted .xlsx. | Show how to configure Aspose.Cells workbook.Settings to enable strong encryption with a custom 256‑bit password before calling Save in a .NET application.
// Common Searches: C# Aspose.Cells how to apply 256‑bit password protection to an XLSX file | Using Microsoft Strong Cryptographic Provider with Aspose.Cells for Excel encryption | Save encrypted workbook with custom password in Aspose.Cells .NET | Set workbook.Settings.Password for strong encryption in Aspose.Cells
// Tags: Aspose.Cells strong encryption Microsoft provider | C# set 256‑bit password XLSX Aspose.Cells | encrypt Excel workbook Aspose.Cells .NET | workbook.Settings.Password strong encryption | save encrypted .xlsx with Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a workbook, adds sample data, assigns a 256‑bit password via workbook.Settings.Password, and saves it as an encrypted XLSX using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some sample data
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Apply password protection to the workbook
            workbook.Settings.Password = "MyStrong256BitPassword!@#123";

            // Save the workbook with the password applied
            workbook.Save("EncryptedWorkbook.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
