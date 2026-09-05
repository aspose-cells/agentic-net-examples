// Title: C# – Detect encrypted XLSX workbook and examine encryption algorithm exposure using Aspose.Cells for .NET
// AI Prompts: Write a C# method that loads an XLSX file with Aspose.Cells LoadOptions and returns true if the workbook is encrypted. | Create C# code that catches a CellsException indicating a required password, logs that the file is encrypted, and notes that the encryption algorithm identifier is not exposed by Aspose.Cells. | Develop a reusable C# utility that accepts a file path, determines whether the Office Open XML workbook is password‑protected with Aspose.Cells, and outputs the encryption algorithm OID when available, otherwise reports that the algorithm cannot be retrieved.
// Common Searches: how to programmatically check if an xlsx file is password protected using Aspose.Cells in C# | Aspose.Cells detect encrypted Excel workbook and retrieve encryption algorithm OID | C# load encrypted XLSX with Aspose.Cells without providing a password | retrieve encryption details of Office Open XML workbook using Aspose.Cells .NET
// Tags: Aspose.Cells detect encrypted XLSX | Aspose.Cells LoadOptions password protection | C# read encryption status of Office Open XML workbook | audit Excel encryption algorithm .NET | handle CellsException for encrypted workbook

using System;
using System.IO;
using Aspose.Cells;

// The sample accepts a file path, verifies the XLSX exists, and attempts to open it with Aspose.Cells LoadOptions. Successful load means the workbook is not encrypted; a CellsException containing "Password" indicates encryption. The code notes that the current Aspose.Cells API does not expose the encryption algorithm identifier, which is useful information for auditing purposes.
class Program
{
    static void Main(string[] args)
    {
        // Expect a file path argument.
        if (args.Length == 0)
        {
            Console.WriteLine("Please provide the path to the XLSX file as an argument.");
            return;
        }

        string filePath = args[0];

        // Ensure the file exists before attempting to load it.
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook without a password.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, the workbook is not encrypted.
            Console.WriteLine("The file is not encrypted.");
        }
        catch (CellsException ex)
        {
            // If the exception message indicates a password is required, treat it as encrypted.
            if (ex.Message != null && ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The file is encrypted.");
                // Encryption details (algorithm, etc.) are not exposed directly in this version of Aspose.Cells.
            }
            else
            {
                // Other CellsException types.
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception e)
        {
            // General exception handling for unexpected errors.
            Console.WriteLine($"Error: {e.Message}");
        }
    }
}
