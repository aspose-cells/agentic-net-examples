// Title: Open a password‑protected Excel workbook with Aspose.Cells LoadOptions in C# and save it without a password
// AI Prompts: Write C# code that creates a LoadOptions object with a password, loads an encrypted .xlsx file using Aspose.Cells, reads the first worksheet name, and saves a new copy without any protection. | Show how to catch a CellsException when the supplied password is incorrect while loading a protected Excel file with Aspose.Cells in C#.
// Common Searches: Aspose.Cells C# load encrypted Excel file with password | how to remove password from .xlsx using Aspose.Cells | C# LoadOptions Password property example Aspose.Cells | open password protected workbook and save unprotected Aspose.Cells | catch CellsException for wrong password Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions password | C# open encrypted .xlsx Aspose.Cells | remove Excel workbook password Aspose.Cells | save unprotected workbook Aspose.Cells | handle CellsException invalid password

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the presence of a password‑protected Excel file, sets the password in a LoadOptions object, loads the workbook with Aspose.Cells, prints the first worksheet name, and then saves a copy of the workbook without any password protection.
class Program
{
    static void Main()
    {
        // Path to the password‑protected Excel file
        string inputPath = "protected.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Create LoadOptions and set the password required to open the file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "myPassword"
            };

            // Load the workbook using the LoadOptions with the correct password
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example operation: write the name of the first worksheet to console
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"First worksheet name: {firstSheet.Name}");

            // Optional: save the workbook to a new file (e.g., without password)
            string outputPath = "unprotected.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved without password to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handles errors related to Aspose.Cells operations (e.g., invalid password)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles any other unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
