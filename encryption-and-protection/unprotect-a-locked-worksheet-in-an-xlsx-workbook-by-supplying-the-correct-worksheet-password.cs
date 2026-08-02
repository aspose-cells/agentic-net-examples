// Title: Unprotect a worksheet in an XLSX file with Aspose.Cells for .NET
// Description: Loads a protected XLSX workbook, removes the worksheet password using Worksheet.Unprotect, verifies the IsProtected flag, and saves an unprotected copy.
// Keywords: Aspose.Cells | C# worksheet unprotect | Worksheet.Unprotect | remove Excel sheet password | check IsProtected | save unprotected workbook
// Common Searches: Aspose.Cells unprotect worksheet C# | How to remove password from Excel sheet using Aspose.Cells | Worksheet.Unprotect example .NET | Check if worksheet is protected Aspose.Cells | Save workbook after unprotecting sheet C#
// Developer Intent: Remove password protection from a specific worksheet in an existing XLSX workbook using Aspose.Cells for .NET.
// Use Cases: Load a workbook, call Worksheet.Unprotect with the known password, and save the file so the sheet can be edited. | Validate that the sheet is no longer protected by reading the IsProtected property before performing data extraction or modification. | Automate batch processing of multiple workbooks, unprotecting designated sheets and generating unprotected versions for downstream analytics.
// AI Prompts: Generate C# code that uses Aspose.Cells to unprotect the second worksheet of a workbook given a password and saves the result to a new file. | Create error‑handling logic for worksheet unprotection when the password may be incorrect, including verification of the IsProtected flag after the call. | Write a reusable method that accepts a file path and password, unprotects all worksheets in the workbook, and returns the names of sheets that were successfully unprotected.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a protected XLSX workbook, removes the worksheet password using Worksheet.Unprotect, verifies the IsProtected flag, and saves an unprotected copy.
    public class UnprotectWorksheetDemo
    {
        public static void Run()
        {
            string inputPath = "ProtectedWorksheet.xlsx";
            string outputPath = "UnprotectedWorksheet.xlsx";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (worksheet is protected, not the file)
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Unprotect the worksheet using the known password
                worksheet.Unprotect("myWorksheetPassword");

                // Verify that the worksheet is no longer protected
                Console.WriteLine("Worksheet IsProtected: " + worksheet.IsProtected);

                // Save the workbook with the worksheet now unprotected
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            UnprotectWorksheetDemo.Run();
        }
    }
}
