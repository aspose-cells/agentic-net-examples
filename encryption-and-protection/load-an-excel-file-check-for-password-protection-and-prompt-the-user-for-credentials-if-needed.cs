// Title: Load an Excel workbook in C# with Aspose.Cells and prompt for a password when the file is protected
// Common Searches: C# Aspose.Cells load encrypted xlsx file with user password prompt | How to open a password protected Excel workbook using Aspose.Cells in .NET | Catch CellsException for protected workbook and request password with Aspose.Cells | Use LoadOptions Password property to open protected Excel files in C# | Check if an Excel file is password protected before loading with Aspose.Cells | Prompt console for Excel password when opening workbook via Aspose.Cells
// Tags: Aspose.Cells load password protected workbook | C# LoadOptions Password property | detect encrypted Excel file Aspose.Cells | handle CellsException protected workbook | console password prompt for Excel file

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the Excel file's existence, attempts to open it with Aspose.Cells without a password, catches a CellsException for protected files, prompts the user for the password via the console, reloads the workbook using LoadOptions with the supplied password, and finally reports the worksheet count or any loading errors.
class Program
{
    static void Main()
    {
        // Path to the Excel file (adjust as needed)
        string filePath = "sample.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(filePath);
        }
        catch (CellsException)
        {
            // Assume the exception is due to a password-protected file
            Console.Write("Enter password for the Excel file: ");
            string password = Console.ReadLine() ?? string.Empty;

            try
            {
                // Load the workbook using the provided password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load the workbook with the supplied password: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load the workbook: {ex.Message}");
            return;
        }

        // At this point the workbook is loaded (either without password or with the supplied one)
        Console.WriteLine($"Workbook loaded successfully. Number of worksheets: {workbook?.Worksheets?.Count ?? 0}");
        // Further processing can be added here
    }
}
