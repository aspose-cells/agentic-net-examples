// Title: Detect if an Excel workbook loaded from a MemoryStream is encrypted using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads an XLSX file into a MemoryStream, attempts to open it with Aspose.Cells Workbook, and returns a Boolean indicating encryption without providing a password. | Show how to wrap the Workbook constructor in a try‑catch block that catches CellsException to identify password‑protected files. | Explain how to use LoadOptions with default settings to test workbook encryption status from a byte array.
// Common Searches: Aspose.Cells C# check workbook encryption from byte array | how to know if an Excel file is password protected when loading from MemoryStream | detect encrypted XLSX using Aspose.Cells without password | C# determine if workbook requires a password before opening with Aspose.Cells | catch CellsException to identify protected Excel files in .NET
// Tags: detect encrypted workbook Aspose.Cells | load Excel from MemoryStream Aspose.Cells | catch CellsException for encryption detection | password‑protected Excel detection C# | Workbook encryption check without password

using System;
using System.IO;
using Aspose.Cells;

// The example reads an XLSX file into a byte array, creates a MemoryStream, and attempts to load it with Aspose.Cells Workbook using default LoadOptions. A try‑catch block captures CellsException, which indicates the workbook is password‑protected, allowing the program to report the encryption status as a Boolean.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook bytes into a memory stream
            byte[] workbookBytes = File.ReadAllBytes(inputPath);
            using (MemoryStream memoryStream = new MemoryStream(workbookBytes))
            {
                // Load options without specifying a password
                LoadOptions loadOptions = new LoadOptions();

                bool isEncrypted = false;
                Workbook workbook = null;

                try
                {
                    // Attempt to load the workbook; if it succeeds, the file is not encrypted
                    workbook = new Workbook(memoryStream, loadOptions);
                    isEncrypted = false;
                }
                catch (CellsException)
                {
                    // Loading failed – most likely because the workbook is password protected
                    isEncrypted = true;
                }
                catch (Exception ex)
                {
                    // Unexpected error while loading the workbook
                    Console.WriteLine($"An unexpected error occurred: {ex.Message}");
                    return;
                }

                Console.WriteLine($"Workbook encrypted: {isEncrypted}");
            }
        }
        catch (Exception ex)
        {
            // General exception handling for file I/O or other runtime errors
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
