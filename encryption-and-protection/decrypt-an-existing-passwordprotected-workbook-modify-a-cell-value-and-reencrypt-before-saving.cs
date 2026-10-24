// Title: Decrypt a password‑protected Excel workbook, modify a cell, and re‑encrypt with a new password using Aspose.Cells for .NET
// Common Searches: C# load an encrypted .xlsx file with Aspose.Cells using the original password | update a cell value in a password‑protected Excel workbook with Aspose.Cells | save an Excel workbook with a new password after editing using Aspose.Cells for .NET | change workbook encryption password programmatically in C# Aspose.Cells | handle CellsException when opening an encrypted Excel file with Aspose.Cells
// Tags: load encrypted workbook with password Aspose.Cells | modify cell value in encrypted Excel using C# | set workbook password before saving Aspose.Cells | save workbook with encryption .NET | handle incorrect password exception Aspose.Cells | re‑encrypt Excel file with new password Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads a password‑protected XLSX file using LoadOptions, updates cell A1 on the first worksheet, assigns a new password via Workbook.Settings, and saves the workbook as a newly encrypted file, handling potential password‑related exceptions.
class Program
{
    static void Main()
    {
        // Paths and passwords
        string inputPath = "encrypted.xlsx";
        string openPassword = "oldPassword";
        string newPassword = "newPassword";
        string outputPath = "encrypted_modified.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook with the original password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = openPassword
            };
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Modify a cell value (example: Sheet1!A1)
            Worksheet sheet = workbook.Worksheets[0];
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Modified value");

            // Set a new password for the workbook (this will encrypt it on save)
            workbook.Settings.Password = newPassword;

            // Save the workbook with the new password
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (CellsException ex)
        {
            // Handles errors related to Aspose.Cells, such as incorrect password
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handles all other runtime errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
