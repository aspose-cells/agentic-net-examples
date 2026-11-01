// Title: Remove opening and worksheet passwords from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove password protection from an Excel workbook with Aspose.Cells | How to unprotect all sheets in a protected XLSX using Aspose.Cells .NET | Aspose.Cells load password‑protected workbook and save without password | Programmatically clear opening password of an XLSX file in C# | Remove worksheet protection with Aspose.Cells without knowing the password
// Tags: Aspose.Cells workbook password removal C# | Aspose.Cells worksheet unprotect method .NET | LoadOptions password protected XLSX Aspose | SaveFormat Xlsx unprotected output Aspose.Cells | Clear workbook Settings.Password Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookUnprotectExample
{
    // The example loads a password‑protected XLSX file using LoadOptions (optionally supplying the opening password), clears the workbook's opening password via Settings.Password, iterates through each worksheet calling Unprotect to drop any sheet passwords, and saves the workbook as an unprotected file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for input (protected) and output (unprotected) workbooks
            string inputFile = "protected.xlsx";
            string outputFile = "unprotected.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            try
            {
                // If the workbook is password‑protected for opening, provide the password here.
                // Leave empty if you don't know the password; the catch block will handle the error.
                string openingPassword = ""; // TODO: set the actual password if known

                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                if (!string.IsNullOrEmpty(openingPassword))
                {
                    loadOptions.Password = openingPassword;
                }

                // Load the workbook (with or without password)
                Workbook workbook = new Workbook(inputFile, loadOptions);

                // Remove the opening password (if any)
                workbook.Settings.Password = string.Empty;

                // Remove protection from each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // If the sheet has a password, passing null/empty removes it
                    sheet.Unprotect();
                }

                // Save the workbook without any passwords
                workbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook has been saved unprotected to \"{outputFile}\".");
            }
            catch (CellsException ex)
            {
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
