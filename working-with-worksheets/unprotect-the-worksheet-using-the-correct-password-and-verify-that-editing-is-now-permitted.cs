// Title: Unprotect a password‑protected worksheet in an Excel file using Aspose.Cells for .NET and verify cell editing
// AI Prompts: Load a workbook from a protected .xlsx file, call Worksheet.Unprotect with the known password, write a value to cell A1 to confirm the sheet is editable, then save the workbook as a new unprotected file using Aspose.Cells in C#. | Open a password‑protected Excel workbook, remove protection from the first worksheet, attempt to modify a cell to ensure protection was removed, and write the resulting unprotected workbook to disk with Aspose.Cells.
// Common Searches: asp.net unprotect worksheet password Aspose.Cells example | c# remove worksheet protection and test editing with Aspose.Cells | how to verify worksheet is unprotected after calling Unprotect in Aspose.Cells | save an unprotected copy of a protected Excel file using Aspose.Cells .NET
// Tags: Aspose.Cells worksheet unprotect password | C# remove Excel sheet protection Aspose.Cells | validate cell edit after worksheet unprotect | save unprotected workbook Aspose.Cells | load protected .xlsx Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads a password‑protected Excel workbook, removes protection from the first worksheet using Worksheet.Unprotect, attempts to write a value to cell A1 to confirm editing is allowed, and saves the result as a new unprotected file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "protected.xlsx";
            const string outputPath = "unprotected.xlsx";
            const string password = "myPassword";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains the protected worksheet
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unprotect the worksheet using the known password
            worksheet.Unprotect(password);

            // Attempt to edit a cell to confirm editing is permitted
            try
            {
                worksheet.Cells["A1"].PutValue("Edited after unprotect");
                Console.WriteLine("Cell edit succeeded.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Cell edit failed: " + ex.Message);
            }

            // Save the workbook with the unprotected worksheet
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
