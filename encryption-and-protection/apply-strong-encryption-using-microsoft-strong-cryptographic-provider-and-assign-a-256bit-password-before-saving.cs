// Title: Encrypt an Excel workbook with a 256‑bit password using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells encrypt Excel file with 256‑bit password | Set password protection on .xlsx using Aspose.Cells .NET API | Apply Microsoft Strong Cryptographic Provider when saving workbook with Aspose.Cells | How to use workbook.Settings.Password for strong encryption in Aspose.Cells | Save encrypted workbook as .xlsx with Aspose.Cells C# example
// Tags: Aspose.Cells workbook password protection | C# encrypt Excel .xlsx with Aspose.Cells | 256-bit password encryption Aspose.Cells | Microsoft Strong Cryptographic Provider Aspose.Cells | Save encrypted workbook using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The sample creates a new Workbook, writes data to the first worksheet, assigns a strong 256‑bit password via workbook.Settings.Password, and saves the file as an encrypted .xlsx using Aspose.Cells for .NET, with exception handling.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Set a password to encrypt the workbook (standard strong encryption)
            workbook.Settings.Password = "StrongPassword256Bit!";

            // Define the output file path
            string outputPath = "EncryptedWorkbook.xlsx";

            // Save the workbook with encryption
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
