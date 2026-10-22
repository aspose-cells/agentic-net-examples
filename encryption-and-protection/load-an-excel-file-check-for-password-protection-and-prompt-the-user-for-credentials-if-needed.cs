// Title: C# example: Open an Excel workbook with Aspose.Cells, detect password protection, and prompt the user for the password
// Common Searches: how to open a password‑protected .xlsx file with Aspose.Cells in C# | using LoadOptions to supply password for an encrypted Excel workbook in Aspose.Cells | catch CellsException for password errors when loading an Excel file in C# | prompt console for Excel workbook password and reload with Aspose.Cells | verify Excel file exists before loading with Aspose.Cells in C#
// Tags: Aspose.Cells load encrypted workbook | C# LoadOptions Password property | detect password protection Aspose.Cells | CellsException password handling | console password prompt Aspose.Cells | file existence check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks whether the specified .xlsx file exists, attempts to load it with Aspose.Cells without a password, catches a password‑related CellsException, prompts the user for the workbook password via the console, reloads the workbook using LoadOptions.Password, handles incorrect passwords and other errors, and finally enumerates the worksheet names when loading succeeds.
class Program
{
    static void Main()
    {
        // Path to the Excel file
        string filePath = "sample.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        Workbook workbook = null;

        // Attempt to load the workbook without a password
        try
        {
            workbook = new Workbook(filePath);
            Console.WriteLine("Workbook loaded successfully without a password.");
        }
        catch (CellsException ex) when (ex.Message.Contains("Password"))
        {
            // The file is password protected; prompt the user for the password
            Console.Write("The workbook is password protected. Please enter the password: ");
            string password = Console.ReadLine();

            // Load the workbook using the supplied password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            try
            {
                workbook = new Workbook(filePath, loadOptions);
                Console.WriteLine("Workbook loaded successfully with the provided password.");
            }
            catch (CellsException innerEx) when (innerEx.Message.Contains("Password"))
            {
                Console.WriteLine("Incorrect password. Unable to open the workbook.");
                return;
            }
            catch (Exception innerEx)
            {
                Console.WriteLine($"An error occurred while loading with password: {innerEx.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            // Handle other possible exceptions (e.g., corrupted file)
            Console.WriteLine($"An error occurred: {ex.Message}");
            return;
        }

        // Ensure the workbook was loaded before proceeding
        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        // Example: list worksheet names
        Console.WriteLine("Worksheets in the workbook:");
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Console.WriteLine($"- {sheet.Name}");
        }
    }
}
