// Title: Prompt user for password and open a protected Excel workbook with Aspose.Cells in C#
// AI Prompts: Write C# code that asks the user for an Excel file path and password, then loads the encrypted workbook using Aspose.Cells LoadOptions. | Show how to handle missing file and incorrect password errors when opening a password‑protected .xlsx with Aspose.Cells.
// Common Searches: C# how to read a password protected .xlsx file using Aspose.Cells LoadOptions | Aspose.Cells open encrypted Excel workbook after prompting user for password | example code to validate Excel file path before loading with Aspose.Cells in .NET | retrieve first worksheet name from a protected workbook using Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions.Password for encrypted XLSX | C# collect workbook path and decryption key | open password protected workbook with Aspose.Cells | extract first sheet title from encrypted workbook | manage absent Excel file errors in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The console app asks for the workbook location and password, verifies the file exists, applies the password via LoadOptions, opens the protected Excel file with Aspose.Cells, and prints the name of the first worksheet while handling missing‑file and password errors.
class Program
{
    static void Main()
    {
        try
        {
            // Prompt for workbook path
            Console.Write("Enter the full path to the workbook: ");
            string filePath = Console.ReadLine()?.Trim();

            // Validate the path
            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("No path was entered. Exiting.");
                return;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Prompt for password
            Console.Write("Enter the password to open the workbook (leave empty if none): ");
            string password = Console.ReadLine();

            // Set load options with password (if any)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            if (!string.IsNullOrEmpty(password))
                loadOptions.Password = password;

            // Load the workbook
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened successfully.");

            // Display first worksheet name, if present
            if (workbook.Worksheets.Count > 0)
                Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
            else
                Console.WriteLine("The workbook contains no worksheets.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
