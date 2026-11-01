// Title: Check if an Excel workbook loaded from a MemoryStream is password‑protected using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells detect password protection when opening workbook from MemoryStream | how to determine if an Excel file is encrypted before loading with Aspose.Cells | using Aspose.Cells to test if workbook requires a password without providing one | catch CellsException to identify encrypted Excel workbook in .NET | verify encryption status of workbook read into MemoryStream with Aspose.Cells
// Tags: detect encrypted workbook Aspose.Cells .NET | memory stream workbook password check | handle CellsException for encrypted Excel | load Excel without password Aspose.Cells | determine workbook encryption status C#

using System;
using System.IO;
using Aspose.Cells;

// The example reads an .xlsx file into a MemoryStream, attempts to open it with Aspose.Cells, catches a CellsException, inspects the exception message for a password hint, and reports whether the workbook is encrypted.
class Program
{
    static void Main()
    {
        // Obtain the workbook data as a memory stream.
        MemoryStream memoryStream = GetWorkbookMemoryStream();
        if (memoryStream == null)
        {
            // Stream could not be created (e.g., file missing).
            return;
        }

        bool isEncrypted = false;

        try
        {
            // Attempt to load the workbook without a password.
            // If the workbook is encrypted, Aspose.Cells will throw a CellsException.
            Workbook workbook = new Workbook(memoryStream);
            // Loaded successfully – assume it is not encrypted.
            isEncrypted = false;
        }
        catch (CellsException ex)
        {
            // Check if the exception indicates a password is required.
            // Different Aspose.Cells versions may have different messages,
            // so we perform a case‑insensitive search for the word "password".
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                isEncrypted = true;
            }
            else
            {
                // Unexpected CellsException – rethrow to surface the error.
                Console.WriteLine($"CellsException: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            // Any other exception (e.g., I/O errors).
            Console.WriteLine($"Error: {ex.Message}");
            return;
        }

        Console.WriteLine(isEncrypted
            ? "The workbook loaded from the memory stream is encrypted."
            : "The workbook loaded from the memory stream is not encrypted.");
    }

    // Reads a workbook file into a memory stream after verifying its existence.
    static MemoryStream GetWorkbookMemoryStream()
    {
        const string filePath = "example.xlsx";

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return null;
        }

        try
        {
            byte[] data = File.ReadAllBytes(filePath);
            return new MemoryStream(data);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read file '{filePath}': {ex.Message}");
            return null;
        }
    }
}
