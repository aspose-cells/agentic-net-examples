using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Path to the OOXML file to be inspected
        string filePath = "encrypted.xlsx"; // TODO: replace with actual file path

        // Verify that the file exists before attempting any operation
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Detect the file format and whether it is encrypted
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);
            Console.WriteLine($"Is Encrypted: {formatInfo.IsEncrypted}");

            if (formatInfo.IsEncrypted)
            {
                // Prompt for password (replace with known password if desired)
                Console.Write("Enter password for the encrypted workbook: ");
                string password = Console.ReadLine();

                // Load the workbook using the supplied password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto)
                {
                    Password = password
                };

                Workbook workbook = new Workbook(filePath, loadOptions);

                // Verify that the workbook reports being encrypted
                Console.WriteLine($"Workbook Settings IsEncrypted: {workbook.Settings.IsEncrypted}");

                // Note: Aspose.Cells does not expose the encryption algorithm identifier directly.
                // For auditing purposes, you would need to parse the encrypted package manually
                // or rely on external tools to retrieve the algorithm name.
            }
        }
        catch (CellsException ex)
        {
            // Handles errors thrown by Aspose.Cells (e.g., invalid password)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}