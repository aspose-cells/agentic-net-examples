using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the text file that contains Excel file paths (one per line)
            string txtFilePath = "excel_files.txt";

            if (!File.Exists(txtFilePath))
            {
                Console.WriteLine($"Text file not found: {txtFilePath}");
                return;
            }

            // Read each line (each should be a path to an Excel workbook)
            foreach (string line in File.ReadLines(txtFilePath))
            {
                string excelPath = line.Trim();

                if (string.IsNullOrEmpty(excelPath))
                    continue; // skip empty lines

                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"Excel file not found: {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook (macro-enabled formats such as .xlsm are supported)
                    Workbook workbook = new Workbook(excelPath);

                    // Access the VBA project associated with the workbook
                    VbaProject vbaProject = workbook.VbaProject;

                    // Determine protection status
                    bool isProtected = vbaProject.IsProtected;
                    bool isLockedForViewing = vbaProject.IslockedForViewing;

                    // Output the results
                    Console.WriteLine($"File: {excelPath}");
                    Console.WriteLine($"  VBA Project Protected: {isProtected}");
                    Console.WriteLine($"  VBA Project Locked for Viewing: {isLockedForViewing}");
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                }
            }
        }
    }
}