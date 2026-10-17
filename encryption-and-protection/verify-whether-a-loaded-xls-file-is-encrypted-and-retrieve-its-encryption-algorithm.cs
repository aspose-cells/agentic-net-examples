// Title: Determine whether an XLS workbook is encrypted using Aspose.Cells for .NET
// Common Searches: how to check if an XLS file is encrypted with Aspose.Cells in C# | retrieve encryption algorithm of a password‑protected Excel workbook using Aspose.Cells .NET | detect encrypted XLS workbook without providing a password in C# code | Aspose.Cells catch CellsException to identify encryption type of an Excel file | programmatically verify password protection of an XLS workbook with Aspose.Cells
// Tags: Aspose.Cells detect workbook encryption | XLS encryption detection Aspose.Cells | retrieve Excel encryption algorithm .NET | handle CellsException encrypted workbook | check password protection Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an XLS file with Aspose.Cells Workbook inside a try‑catch block. If loading succeeds, the workbook is not encrypted. If a CellsException is thrown and its message contains the word "password", the code reports the workbook as encrypted. Other exceptions are also handled, providing a simple way to verify encryption status of an XLS file in C#.
class Program
{
    static void Main()
    {
        // Path to the XLS file to be examined
        string filePath = "input.xls";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook without a password
            Workbook workbook = new Workbook(filePath);
            // If loading succeeds, the workbook is not encrypted
            Console.WriteLine("Workbook is not encrypted.");
        }
        catch (CellsException ex)
        {
            // If the exception message indicates a password is required, the workbook is encrypted
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Workbook is encrypted.");
            }
            else
            {
                // Handle other CellsException scenarios
                Console.WriteLine($"CellsException occurred: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
