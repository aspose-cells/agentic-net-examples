using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled) to be checked
            string excelPath = "example.xlsm";

            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: File '{excelPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed and if the signature is valid
            bool isSigned = vbaProject.IsSigned;
            bool isValidSigned = vbaProject.IsValidSigned;

            // Prepare CSV output
            string csvPath = "VbaSignatureReport.csv";
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                // Write header
                writer.WriteLine("FileName,IsSigned,IsValidSigned");

                // Write data row
                writer.WriteLine($"{Path.GetFileName(excelPath)},{isSigned},{isValidSigned}");
            }

            Console.WriteLine($"VBA signature check completed. Report saved to '{csvPath}'.");
        }
    }
}