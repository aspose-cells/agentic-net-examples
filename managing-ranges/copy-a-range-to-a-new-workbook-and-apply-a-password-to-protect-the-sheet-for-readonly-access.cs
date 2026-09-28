// Title: Copy a specific range from an existing workbook to a new workbook and apply password protection to the sheet using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to copy cells A1:C10 from source.xlsx into a new workbook named destination.xlsx while preserving formulas and formatting. | Create a new workbook without the default sheet, add a worksheet called "CopiedRange", and paste the copied range at the top‑left corner. | Apply read‑only password protection to the newly added worksheet using the Protect method with ProtectionType.All.
// Common Searches: Aspose.Cells copy range to new workbook C# example | How to protect an Aspose.Cells worksheet with a password for read‑only access | Remove default worksheet before adding a custom sheet in Aspose.Cells | Copy cells with formulas and styles using Aspose.Range in C#
// Tags: copy cell block to separate workbook Aspose.Cells | read‑only worksheet protection Aspose.Cells C# | clear default worksheet Aspose.Cells before adding new | Aspose.Range transfer formulas styles | sheet protection using ProtectionType.All Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The sample loads source.xlsx, creates a new workbook without the default sheet, adds a worksheet named "CopiedRange", copies the A1:C10 range (including values, formulas, and styles) from the source sheet to the new sheet, protects the worksheet with a read‑only password, and saves the result as destination.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (e.g., A1:C10)
            int startRow = 0;          // Row index for A1 (zero‑based)
            int startColumn = 0;       // Column index for A1 (zero‑based)
            int totalRows = 10;        // Number of rows to copy
            int totalColumns = 3;      // Number of columns to copy (A‑C)

            // Create a new workbook for the copied range
            Workbook destinationWorkbook = new Workbook();
            destinationWorkbook.Worksheets.Clear(); // Remove default sheet
            Worksheet destinationSheet = destinationWorkbook.Worksheets.Add("CopiedRange");

            // Copy the defined range from the source sheet to the destination sheet
            // Using AsposeRange.Copy to transfer values, formulas, and styles
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
            AsposeRange destRange = destinationSheet.Cells.CreateRange(0, 0, totalRows, totalColumns);
            sourceRange.Copy(destRange);

            // Apply password protection to the destination sheet for read‑only access
            string password = "myPassword";
            destinationSheet.Protect(ProtectionType.All, password, string.Empty);

            // Save the new workbook
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Range copied and saved to \"{destinationPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
