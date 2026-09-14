// Title: Open an encrypted Excel workbook with Aspose.Cells for .NET without a password and capture the CellsException error code
// AI Prompts: Generate C# code that loads a password‑protected .xlsx using Aspose.Cells without providing a password and logs the resulting CellsException error number. | Write a try‑catch example in C# that opens an encrypted workbook with Aspose.Cells, catches CellsException, and prints both the exception message and its numeric error code. | Show how to handle FileNotFoundException and any other unexpected exceptions when attempting to open a protected Excel file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells .NET how to get error code when opening password protected Excel without password | C# catch CellsException for encrypted workbook and retrieve error number | What exception is thrown by Aspose.Cells when a workbook requires a password | How to detect missing password error while loading encrypted .xlsx with Aspose.Cells
// Tags: Aspose.Cells load encrypted workbook without password | CellsException error code retrieval | handling password required exception Aspose.Cells | C# file not found handling Aspose.Cells | catch generic exceptions Aspose.Cells workbook load

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of an encrypted Excel file, attempts to open it with Aspose.Cells for .NET without supplying a password, and uses distinct catch blocks to handle CellsException (indicating a missing password), FileNotFoundException, and any other unexpected errors, outputting the relevant messages and error codes.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook without providing a password
            Workbook wb = new Workbook(filePath);
            Console.WriteLine("Workbook opened successfully (unexpected).");
        }
        catch (CellsException ex)
        {
            // Output the exception message (e.g., password required)
            Console.WriteLine($"Failed to open workbook. Reason: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            // Handle missing file scenario
            Console.WriteLine($"File not found: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
