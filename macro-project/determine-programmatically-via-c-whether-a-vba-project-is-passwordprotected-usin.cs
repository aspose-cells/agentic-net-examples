using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace VbaProtectionChecker
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the text file that contains the Excel file path (first line)
            string txtFilePath = "input.txt";

            if (!File.Exists(txtFilePath))
            {
                Console.WriteLine($"Text file not found: {txtFilePath}");
                return;
            }

            // Read the first non‑empty line from the text file
            string excelFilePath = null;
            foreach (var line in File.ReadLines(txtFilePath))
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    excelFilePath = line.Trim();
                    break;
                }
            }

            if (string.IsNullOrEmpty(excelFilePath) || !File.Exists(excelFilePath))
            {
                Console.WriteLine("Excel file path is missing or the file does not exist.");
                return;
            }

            // Load the workbook (macro‑enabled formats such as .xlsm, .xlsb, etc.)
            Workbook workbook = new Workbook(excelFilePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is password‑protected
            bool isProtected = vbaProject.IsProtected;

            // Output the result
            Console.WriteLine($"Workbook: {Path.GetFileName(excelFilePath)}");
            Console.WriteLine($"VBA Project Password Protected: {isProtected}");
        }
    }
}