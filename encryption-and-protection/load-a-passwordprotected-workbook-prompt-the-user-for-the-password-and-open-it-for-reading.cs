// Title: Load a password‑protected Excel workbook in C# by prompting the user for the password with Aspose.Cells
// Common Searches: C# Aspose.Cells open password protected Excel file from console input | How to use LoadOptions to read an encrypted workbook with a user‑supplied password | Prompt user for Excel file password then load workbook using Aspose.Cells .NET | Read first worksheet name from a password‑protected Excel workbook in a C# console app | Validate file path and open protected Excel workbook with Aspose.Cells LoadOptions
// Tags: Aspose.Cells LoadOptions password protection | open encrypted Excel workbook C# | read password‑protected workbook Aspose.Cells | console password prompt for Excel file C# | file existence check before loading workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The console program asks for the Excel file path and password, verifies the file exists, sets the password in a LoadOptions object, loads the protected workbook with Aspose.Cells, and displays the name of the first worksheet if present.
class Program
{
    static void Main()
    {
        try
        {
            // Prompt the user for the workbook file path
            Console.Write("Enter the full path to the workbook: ");
            string filePath = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                Console.WriteLine("File path cannot be empty.");
                return;
            }

            // Verify that the file exists to avoid FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Prompt the user for the password
            Console.Write("Enter the password to open the workbook: ");
            string password = Console.ReadLine() ?? string.Empty;

            // Configure load options with the supplied password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the password‑protected workbook for reading
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Example operation: display the name of the first worksheet
            Console.WriteLine("Workbook opened successfully.");
            if (workbook.Worksheets.Count > 0)
            {
                Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);
            }
            else
            {
                Console.WriteLine("The workbook contains no worksheets.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
