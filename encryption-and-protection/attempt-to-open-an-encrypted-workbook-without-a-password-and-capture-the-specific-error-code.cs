// Title: Open an encrypted Excel workbook without a password using Aspose.Cells for .NET and capture the specific CellsException error code
// Common Searches: Aspose.Cells .NET load encrypted Excel file without password and retrieve error code | How to catch CellsException when opening a password‑protected workbook in C# | Retrieve specific error number for encrypted workbook load failure using Aspose.Cells | C# example of handling missing password error with Aspose.Cells LoadOptions | What exception is thrown by Aspose.Cells when opening a protected XLSX without a password
// Tags: Aspose.Cells load encrypted workbook without password | CellsException error code password protection | C# handling encrypted Excel file Aspose.Cells | LoadOptions usage for protected XLSX | Aspose.Cells workbook opening error handling

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the presence of an encrypted Excel file, creates a LoadOptions object without a password, attempts to open the workbook with Aspose.Cells, and catches the CellsException to display the specific error message and code, while also handling any unexpected exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "EncryptedWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // LoadOptions without specifying a password
            LoadOptions loadOptions = new LoadOptions();

            // Attempt to open the encrypted workbook
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened successfully.");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors
            Console.WriteLine($"Failed to open workbook. Details: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
