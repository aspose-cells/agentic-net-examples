using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Expect two arguments: input workbook path and output CSV path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: AsposeCellsVbaSignatureValidation <inputWorkbook.xlsm> <output.csv>");
                return;
            }

            string inputPath = args[0];
            string csvPath = args[1];

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed and if the signature is valid
            bool isSigned = vbaProject.IsSigned;               // property rule
            bool isValidSigned = vbaProject.IsValidSigned;    // property rule

            // Prepare CSV content
            // Header: FileName,IsSigned,IsValidSigned
            // Data row: input file name, true/false values
            string fileName = Path.GetFileName(inputPath);
            string csvLine = $"{fileName},{isSigned},{isValidSigned}";

            // Write result to CSV file (save rule)
            using (StreamWriter writer = new StreamWriter(csvPath, false))
            {
                writer.WriteLine("FileName,IsSigned,IsValidSigned");
                writer.WriteLine(csvLine);
            }

            Console.WriteLine($"Signature validation completed. Results saved to {csvPath}");
        }
    }
}