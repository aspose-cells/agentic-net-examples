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
            // Adjust the file name/extension as needed (e.g., .xltm for macro-enabled templates)
            string fileName = "template_with_macro.xltm";
            string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);

            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Check if the workbook contains VBA macros
            if (workbook.HasMacro)
            {
                // Access the VBA project
                VbaProject vbaProject = workbook.VbaProject;

                // Determine if the VBA project is signed
                bool isSigned = vbaProject.IsSigned;
                Console.WriteLine($"VBA project is signed: {isSigned}");

                // If signed, verify the signature validity
                if (isSigned)
                {
                    bool isValid = vbaProject.IsValidSigned;
                    Console.WriteLine($"VBA signature is valid: {isValid}");
                }
                else
                {
                    Console.WriteLine("VBA project is not signed; no signature to validate.");
                }
            }
            else
            {
                Console.WriteLine("Workbook does not contain VBA macros.");
            }
        }
    }
}