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
            // Path to the Excel file (macro-enabled workbook)
            string excelPath = "sample.xlsm";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(excelPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed
            bool isSigned = vbaProject.IsSigned;

            // Prepare the output text
            string output = isSigned
                ? "VBA project is signed."
                : "VBA project is NOT signed.";

            // Optionally include validity information if signed
            if (isSigned)
            {
                output += Environment.NewLine + "Signature valid: " + vbaProject.IsValidSigned;
            }

            // Write the result to a TXT file (output in TXT format)
            string txtOutputPath = "VbaSignatureStatus.txt";
            File.WriteAllText(txtOutputPath, output);

            // Also write to console for immediate feedback
            Console.WriteLine(output);
            Console.WriteLine("Result written to: " + txtOutputPath);
        }
    }
}