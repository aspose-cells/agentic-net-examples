using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookUnprotectAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the protected workbook
            string inputPath = "ProtectedWorkbook.xlsx";

            // Recovered password for the workbook structure
            string password = "recoveredPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"[{DateTime.UtcNow:u}] Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook (lifecycle: load)
                Workbook workbook = new Workbook(inputPath);

                // Check if the workbook structure is protected with a password
                if (workbook.IsWorkbookProtectedWithPassword)
                {
                    try
                    {
                        // Attempt to unprotect the workbook using the provided password
                        workbook.Unprotect(password);
                        Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook '{inputPath}' was unprotected using the provided password.");
                    }
                    catch (CellsException)
                    {
                        // Invalid password scenario
                        Console.WriteLine($"[{DateTime.UtcNow:u}] Invalid password provided for workbook '{inputPath}'.");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine($"[{DateTime.UtcNow:u}] Workbook '{inputPath}' is not protected with a password.");
                }

                // Save the unprotected workbook (lifecycle: save)
                string outputPath = "UnprotectedWorkbook.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved as '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"[{DateTime.UtcNow:u}] An error occurred: {ex.Message}");
            }
        }
    }
}