using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class UnprotectWorkbookStructure
    {
        public static void Run()
        {
            const string inputPath = "protected_workbook.xlsx";
            const string outputPath = "unprotected_workbook.xlsx";
            const string password = "yourPassword";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the encrypted workbook using the password
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = password
                };
                var workbook = new Workbook(inputPath, loadOptions);

                // Unprotect the workbook structure (if it is protected with a password)
                workbook.Unprotect(password);

                // Save the unprotected workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for demonstration
        public static void Main()
        {
            Run();
        }
    }
}