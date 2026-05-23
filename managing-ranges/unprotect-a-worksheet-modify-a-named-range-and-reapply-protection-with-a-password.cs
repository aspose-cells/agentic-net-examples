using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "InputWorkbook.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Unprotect the worksheet (use empty string if no password)
                sheet.Unprotect("oldPassword");

                // Retrieve the named range "MyRange"
                Name namedRange = workbook.Worksheets.Names["MyRange"];
                if (namedRange != null)
                {
                    // Get the actual range object (use alias to avoid ambiguity with System.Range)
                    AsposeRange range = namedRange.GetRange();

                    // Example modification: set each cell in the range to "Updated"
                    foreach (Cell cell in range)
                    {
                        cell.PutValue("Updated");
                    }
                }
                else
                {
                    Console.WriteLine("Named range 'MyRange' not found.");
                }

                // Re‑apply protection with a new password
                sheet.Protect(ProtectionType.All, "newPassword", null);

                // Save the modified workbook
                string outputPath = "OutputWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}