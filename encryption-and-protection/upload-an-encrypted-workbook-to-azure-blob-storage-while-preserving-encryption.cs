using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsAzureUpload
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- Create and encrypt the workbook ----------
                // Create a new workbook (uses Workbook() constructor rule)
                using (Workbook workbook = new Workbook())
                {
                    // Add some sample data
                    Worksheet sheet = workbook.Worksheets[0];
                    sheet.Cells["A1"].PutValue("Encrypted");
                    sheet.Cells["B1"].PutValue(DateTime.Now);

                    // Set a password to protect the workbook (WorkbookSettings.Password property)
                    workbook.Settings.Password = "MySecretPwd";

                    // Optional: set stronger encryption options (SetEncryptionOptions method)
                    workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 128);

                    // ---------- Save the encrypted workbook to a memory stream ----------
                    // Use the Save(Stream, SaveFormat) rule to keep encryption intact
                    using (MemoryStream workbookStream = new MemoryStream())
                    {
                        workbook.Save(workbookStream, SaveFormat.Xlsx);
                        workbookStream.Position = 0; // Reset for reading

                        // ---------- Save the stream to a local file ----------
                        // This replaces Azure upload when the Azure SDK is unavailable.
                        string outputPath = Path.Combine(Environment.CurrentDirectory, "encrypted_workbook.xlsx");

                        // Ensure the directory exists
                        string? directory = Path.GetDirectoryName(outputPath);
                        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }

                        // Write the stream to the file
                        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                        {
                            workbookStream.CopyTo(fileStream);
                        }

                        Console.WriteLine($"Encrypted workbook saved to: {outputPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}