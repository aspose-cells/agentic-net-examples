// Title: How to remove the edit (modify) password from an Excel .xlsx workbook while keeping the opening password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens a password‑protected .xlsx file with its read‑only password, clears any modify password, and saves the workbook preserving the original opening password using Aspose.Cells. | Generate a C# example that loads an Excel workbook via LoadOptions with an opening password, sets workbook.Settings.Password to keep the open password, and ensures no edit protection password is stored when saving with Aspose.Cells. | Provide a step‑by‑step C# snippet that removes the 'password to modify' from a workbook while retaining the existing opening password, leveraging Aspose.Cells' LoadOptions and Workbook.Settings.
// Common Searches: Aspose.Cells C# remove edit password from Excel file but keep open password | how to clear modify protection in .xlsx using Aspose.Cells .NET | load workbook with opening password and delete modify password Aspose.Cells | C# Aspose.Cells preserve read password after removing edit protection
// Tags: Aspose.Cells clear modify password C# | Aspose.Cells preserve opening password .NET | LoadOptions opening password Aspose.Cells | Workbook.Settings.Password edit protection removal | Excel .xlsx edit password removal Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an .xlsx workbook using the opening (read) password via LoadOptions, clears any modify password, reassigns the opening password to workbook.Settings.Password, and saves the file, effectively removing the edit protection while retaining the original opening password.
class Program
{
    static void Main()
    {
        const string inputPath = "InputWorkbook.xlsx";
        const string outputPath = "OutputWorkbook.xlsx";
        const string openingPassword = "OpenPassword123";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook using the opening (read) password.
            var loadOptions = new LoadOptions
            {
                Password = openingPassword
            };
            var workbook = new Workbook(inputPath, loadOptions);

            // Preserve the opening password when saving.
            workbook.Settings.Password = openingPassword;

            // Save the workbook.
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
