// Title: Check if a worksheet is password protected and skip it before editing with Aspose.Cells for .NET
// Common Searches: how to determine if an Excel worksheet is protected using Aspose.Cells in C# | skip password‑protected sheets when updating cells with Aspose.Cells | prompt for workbook password and load with LoadOptions in Aspose.Cells | iterate worksheets and modify only unprotected ones Aspose.Cells C# | check worksheet IsProtected flag before writing data with Aspose.Cells
// Tags: Aspose.Cells detect worksheet protection | C# skip protected worksheets Aspose.Cells | LoadOptions password handling Aspose.Cells | IsProtected property Excel worksheet Aspose.Cells | modify unprotected sheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, asks for a password if the file is encrypted, then iterates through each worksheet, uses the IsProtected flag to identify protected sheets, skips those sheets, writes a value to cell A1 of unprotected sheets, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load without a password
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            // If loading fails due to password protection, prompt for password
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.Write("Workbook is password protected. Enter password: ");
                string password = Console.ReadLine() ?? string.Empty;

                var loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                try
                {
                    workbook = new Workbook(inputPath, loadOptions);
                }
                catch (Exception innerEx)
                {
                    Console.WriteLine($"Failed to load workbook with password: {innerEx.Message}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }
        }

        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded.");
            return;
        }

        try
        {
            // Process each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.IsProtected)
                {
                    Console.WriteLine($"Worksheet '{sheet.Name}' is password protected. Skipping modifications.");
                    continue;
                }

                // Example modification
                sheet.Cells["A1"].PutValue("Modified");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
