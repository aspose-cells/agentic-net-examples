// Title: C# console program to detect if an Excel workbook (.xls/.xlsx) is password‑protected using Aspose.Cells
// AI Prompts: Write a C# console application that accepts a file path argument, loads the workbook with Aspose.Cells LoadOptions (no password), and prints whether the file is encrypted. | Refactor the code to expose a method that returns a boolean indicating encryption status instead of writing messages directly to the console. | Extend the solution to handle both .xls and .xlsx formats, capture non‑password loading errors, and report them separately.
// Common Searches: how to programmatically determine if an Excel file is password protected using Aspose.Cells in .NET | C# detect encrypted .xlsx workbook without providing password | Aspose.Cells LoadOptions exception handling for password‑protected Excel files | console app to check Excel file encryption status and handle load errors | detect password protection for .xls files with Aspose.Cells C#
// Tags: detect encrypted workbook Aspose.Cells | check Excel password protection C# | load workbook without password Aspose.Cells | handle CellsException password required | support xls xlsx encryption detection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The console app receives an Excel file path, attempts to load the workbook with Aspose.Cells using default LoadOptions, and determines encryption by catching a CellsException that contains a password‑related message, then outputs the appropriate status.
class Program
{
    static void Main(string[] args)
    {
        // Verify that a file path argument was supplied
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the Excel file path as a command‑line argument.");
            return;
        }

        string filePath = args[0];

        // Ensure the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook without a password
            LoadOptions loadOptions = new LoadOptions(); // no password supplied
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, the file is not encrypted
            Console.WriteLine("The file is not encrypted.");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws a CellsException when a password is required
            // Check the exception message for password‑related text
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The file is encrypted. Password is required to open it.");
            }
            else
            {
                Console.WriteLine($"Error loading file: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
