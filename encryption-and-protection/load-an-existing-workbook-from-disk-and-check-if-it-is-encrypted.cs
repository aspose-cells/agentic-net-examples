// Title: Determine whether an Excel workbook on disk is password‑protected using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells check if an .xlsx file requires a password before opening | how to detect encrypted Excel workbook with Aspose.Cells without providing a password | using try‑catch to identify password‑protected workbook in Aspose.Cells .NET | programmatically verify Excel file encryption status with Aspose.Cells in C# | detect password requirement when loading workbook with Aspose.Cells API
// Tags: detect encrypted workbook Aspose.Cells | load .xlsx without password C# Aspose.Cells | CellsException handling password protection | programmatic workbook encryption check .NET | Excel file protection verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file from a specified path, attempts to open it with Aspose.Cells without a password, catches a CellsException that mentions "password" to infer encryption, and outputs whether the workbook is password‑protected.
class Program
{
    static void Main()
    {
        // Path to the workbook file on disk
        string filePath = @"C:\Path\To\Your\Workbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        bool isEncrypted = false;
        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without providing a password
            workbook = new Workbook(filePath);
            // If loading succeeds, the workbook is not encrypted
            isEncrypted = false;
        }
        catch (CellsException ex)
        {
            // If a password is required, Aspose.Cells throws a CellsException
            // Treat this as an indication that the workbook is encrypted
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                isEncrypted = true;
            }
            else
            {
                // Other CellsException – report and exit
                Console.WriteLine($"CellsException: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            // Handle any other runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
            return;
        }

        Console.WriteLine($"Workbook encrypted: {isEncrypted}");
    }
}
