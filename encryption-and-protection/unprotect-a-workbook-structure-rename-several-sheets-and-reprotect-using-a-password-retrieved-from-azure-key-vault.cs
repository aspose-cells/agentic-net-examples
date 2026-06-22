using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        try
        {
            // Retrieve the workbook password (e.g., from an environment variable)
            // If no password is set, an empty string is used (works for unprotected workbooks)
            string password = Environment.GetEnvironmentVariable("WorkbookPassword") ?? string.Empty;

            // Input workbook path
            string inputFile = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the workbook with the password (if any)
            var loadOptions = new LoadOptions();
            if (!string.IsNullOrEmpty(password))
                loadOptions.Password = password;

            var workbook = new Workbook(inputFile, loadOptions);

            // Unprotect the workbook structure if a password is provided
            if (!string.IsNullOrEmpty(password))
                workbook.Unprotect(password);

            // Rename the first few worksheets (ensure names are safe for Excel)
            string[] desiredNames = { "Summary", "Data", "Report" };
            int renameCount = Math.Min(desiredNames.Length, workbook.Worksheets.Count);
            for (int i = 0; i < renameCount; i++)
            {
                string safeName = CellsHelper.CreateSafeSheetName(desiredNames[i]);
                workbook.Worksheets[i].Name = safeName;
            }

            // Re‑protect the workbook structure with the same password (if any)
            if (!string.IsNullOrEmpty(password))
                workbook.Protect(ProtectionType.Structure, password);

            // Output workbook path
            string outputFile = "output.xlsx";

            // Save the modified workbook
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}