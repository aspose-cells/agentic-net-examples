// Title: How to protect an Excel worksheet and export it to CSV using Aspose.Cells for .NET without losing data
// Common Searches: Aspose.Cells protect worksheet then save as CSV in C# | Export protected Excel sheet to CSV while keeping protection settings using Aspose.Cells | Temporarily unprotect a worksheet for CSV export with Aspose.Cells .NET | C# code to apply worksheet protection and generate CSV using Aspose.Cells | Handle password‑protected workbook when exporting to CSV with Aspose.Cells
// Tags: worksheet protection API Aspose.Cells | CSV export operation Aspose.Cells | temporary worksheet unprotect Aspose.Cells | password protected workbook handling Aspose.Cells | save workbook as CSV C# Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX file, applies full protection to the first worksheet, temporarily removes the protection to save the sheet as a CSV file, then re‑applies protection. It also validates the input file's existence and demonstrates handling of password‑protected workbooks via exception handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook (no password is supplied; handle password‑protected files via exception)
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Apply full protection to the worksheet
            worksheet.Protect(ProtectionType.All);

            // Temporarily remove protection for CSV export
            worksheet.Unprotect();

            // Export the worksheet data to CSV format
            workbook.Save(outputPath, SaveFormat.Csv);

            // Re‑apply protection after export (if further operations are needed)
            worksheet.Protect(ProtectionType.All);

            Console.WriteLine($"Workbook exported successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handle password‑protected workbook or other CellsException cases
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The workbook is password protected. Provide a valid password in the code.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
