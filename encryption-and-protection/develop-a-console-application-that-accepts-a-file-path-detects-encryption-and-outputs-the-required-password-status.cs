// Title: C# console app to check if an Excel workbook is password‑protected using Aspose.Cells LoadOptions
// Common Searches: how to programmatically determine if an Excel file is encrypted with Aspose.Cells in C# | detect password requirement for .xlsx using Aspose.Cells LoadOptions without opening the workbook | C# console utility to test Excel workbook for password protection using Aspose.Cells | Aspose.Cells CellsException password detection example in .NET | check if Excel workbook needs a password before loading with Aspose.Cells
// Tags: Aspose.Cells LoadOptions password detection | C# detect Excel encryption Aspose | CellsException password required handling | console application Excel protection check | Aspose.Cells workbook encryption detection

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that receives an Excel file path, uses Aspose.Cells LoadOptions without a password to attempt loading, catches a CellsException containing the word "password" to identify encrypted workbooks, and prints either "Password required" or "No password required".
class Program
{
    static void Main(string[] args)
    {
        // Verify that a file path was supplied.
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: <program> <excel-file-path>");
            return;
        }

        string filePath = args[0];

        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found: " + filePath);
            return;
        }

        // LoadOptions without a password – used to detect encryption.
        LoadOptions loadOptions = new LoadOptions();

        Workbook workbook = null;
        try
        {
            // Attempt to load the workbook.
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws a CellsException when a password is required.
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Password required");
            }
            else
            {
                Console.WriteLine($"Error loading file: {ex.Message}");
            }
            return;
        }
        catch (Exception ex)
        {
            // Any other loading error.
            Console.WriteLine($"Error loading file: {ex.Message}");
            return;
        }

        // If loading succeeded, the workbook is not password‑protected.
        Console.WriteLine("No password required");
    }
}
