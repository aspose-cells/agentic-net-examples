// Title: Prompt user for password and load a password‑protected Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: C# code to ask for workbook password and open encrypted Excel file using Aspose.Cells | How to read worksheets from a password‑protected Excel file with Aspose.Cells LoadOptions | Load a protected .xlsx file in .NET by providing password at runtime | Aspose.Cells example for opening Excel workbook with user‑supplied password | Validate file path and open encrypted Excel workbook in C# console application
// Tags: Aspose.Cells LoadOptions password protection | extract worksheet names from protected Excel workbook | validate Excel file existence C# | console user input for workbook password | open .xlsx with password using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The console program asks the user for the full path of a password‑protected Excel file and its password, checks that the file exists, creates a LoadOptions object with the supplied password, loads the workbook via Aspose.Cells, and then displays the total worksheet count and the name of the first sheet, handling any errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            // Prompt the user for the workbook file path
            Console.Write("Enter the full path to the password‑protected workbook: ");
            string filePath = Console.ReadLine()?.Trim();

            // Validate the file path
            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("File path cannot be empty.");
                return;
            }

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Prompt the user for the password
            Console.Write("Enter the password to open the workbook: ");
            string password = Console.ReadLine();

            // Set load options with the provided password (null or empty password is allowed)
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };

            // Load the workbook using the provided password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Workbook loaded successfully
            Console.WriteLine($"Workbook loaded successfully. Number of worksheets: {workbook.Worksheets.Count}");

            // Example: read the name of the first worksheet
            if (workbook.Worksheets.Count > 0)
            {
                Console.WriteLine($"First worksheet name: {workbook.Worksheets[0].Name}");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
        finally
        {
            // Keep console window open
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
}
