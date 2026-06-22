using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook file
            string filePath = "sample.xlsx";

            try
            {
                // Ensure the file exists before processing
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Error: File '{filePath}' not found.");
                    return;
                }

                // Detect format and encryption status without opening the workbook
                FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(filePath);

                if (formatInfo.IsEncrypted)
                {
                    Console.WriteLine("Warning: The workbook is encrypted and requires a password.");

                    // Load the encrypted workbook with a password
                    LoadOptions loadOptions = new LoadOptions { Password = "yourPassword" };
                    Workbook encryptedWorkbook = new Workbook(filePath, loadOptions);
                    Console.WriteLine("Encrypted workbook loaded successfully.");
                }
                else
                {
                    // Load the workbook normally
                    Workbook workbook = new Workbook(filePath);
                    Console.WriteLine("Workbook loaded successfully. It is not encrypted.");
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine($"File not found: {ex.Message}");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}