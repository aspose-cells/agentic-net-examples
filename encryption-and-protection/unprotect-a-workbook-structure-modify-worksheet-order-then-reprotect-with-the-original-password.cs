using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class WorkbookStructureReorder
    {
        static void Main()
        {
            // Paths and password
            string inputPath = "ProtectedWorkbook.xlsx";
            string outputPath = "ReorderedWorkbook.xlsx";
            string password = "myPassword";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the protected workbook
                Workbook workbook = new Workbook(inputPath);

                // Unprotect the workbook (structure and windows) using the password
                workbook.Unprotect(password);

                // Example reordering: move the first worksheet to the end
                if (workbook.Worksheets.Count > 1)
                {
                    Worksheet firstSheet = workbook.Worksheets[0];
                    firstSheet.MoveTo(workbook.Worksheets.Count - 1);
                }

                // Re‑protect only the workbook structure with the same password
                workbook.Protect(ProtectionType.Structure, password);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}