// Title: Save a password‑protected Excel workbook to a MemoryStream and read its length to see encryption overhead with Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to encrypt a workbook with a password, save it directly to a MemoryStream, and output the stream's byte length. | Change the workbook password, re‑save to a MemoryStream, and compare the resulting stream sizes to evaluate encryption impact. | Retrieve the length of an encrypted XLSX file stored in a MemoryStream without writing the file to disk using Aspose.Cells. | Configure OoxmlSaveOptions for password protection and obtain the memory‑stream size after saving the workbook.
// Common Searches: how to get size of password protected Excel file in memory using Aspose.Cells | Aspose.Cells encrypt workbook to memory stream and check length | measure encryption overhead of XLSX with Aspose.Cells .NET | save encrypted workbook to MemoryStream without creating a file Aspose.Cells | retrieve byte count of encrypted Excel workbook in C# Aspose.Cells
// Tags: encrypted workbook memory stream Aspose.Cells | OoxmlSaveOptions password protection .NET | measure encrypted XLSX size in memory | Aspose.Cells workbook encryption overhead | retrieve stream length after workbook save | password protected Excel file size calculation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // For OoxmlSaveOptions

// // Creates a new workbook, adds sample data, applies a password via Settings.Password, saves the encrypted workbook to a MemoryStream using OoxmlSaveOptions, and prints the stream's Length to illustrate the size impact of encryption.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // Add some sample data
            workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");

            // Set password protection for the workbook
            workbook.Settings.Password = "MySecretPassword";

            // Configure save options (no password property needed here)
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the encrypted workbook to a memory stream (lifecycle save rule)
            using (MemoryStream memoryStream = new MemoryStream())
            {
                try
                {
                    workbook.Save(memoryStream, saveOptions);
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error during save: {saveEx.Message}");
                    return;
                }

                // Get the length of the stream to observe encryption overhead
                long encryptedLength = memoryStream.Length;
                Console.WriteLine($"Encrypted workbook stream length: {encryptedLength}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
