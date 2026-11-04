// Title: Handle incorrect password and corrupted file errors separately when opening an encrypted Excel workbook with Aspose.Cells in C#
// Common Searches: Aspose.Cells differentiate wrong password from corrupted encrypted Excel file in C# | C# catch CellsException to identify invalid password versus file corruption | How to detect corrupted encrypted workbook when loading with Aspose.Cells LoadOptions | Error handling for encrypted Excel files using Aspose.Cells password protection | Check file existence before opening encrypted workbook with Aspose.Cells
// Tags: password validation error handling Aspose.Cells | corrupted encrypted workbook detection Aspose.Cells | CellsException message parsing for Excel encryption | load encrypted Excel workbook with LoadOptions C# | file existence verification before Aspose.Cells load

using System;
using System.IO;
using Aspose.Cells;

// Shows how to load an encrypted Excel file with Aspose.Cells in C#, verify the file exists, supply a password via LoadOptions, and use CellsException message analysis to distinguish between an incorrect password and a corrupted or invalid workbook.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        // The password supplied by the user (could be empty or wrong)
        string password = "userSuppliedPassword";

        // LoadOptions with the password to open an encrypted workbook
        LoadOptions loadOptions = new LoadOptions
        {
            Password = password
        };

        try
        {
            // Attempt to load the workbook using the provided password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, proceed with normal processing
            Console.WriteLine("Workbook loaded successfully.");
            // ... (additional processing)
        }
        catch (CellsException ex)
        {
            // Aspose.Cells may not expose an ErrorCode property in some versions.
            // Use the exception message to differentiate common scenarios.
            string msg = ex.Message ?? string.Empty;

            if (msg.IndexOf("incorrect password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The password provided is incorrect.");
                // Prompt user for a new password or abort operation
            }
            else if (msg.IndexOf("corrupt", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     msg.IndexOf("invalid", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The file appears to be corrupted or not a valid Excel file.");
                // Log details, notify user, or attempt recovery steps
            }
            else
            {
                // Re‑throw unexpected exceptions for higher‑level handling
                throw;
            }
        }
        catch (Exception ex)
        {
            // Catch any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
