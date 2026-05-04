using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaProjectSignatureCheckDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the ODS workbook (replace with actual file path)
            string odsFileName = "sample_with_vba.ods";
            string odsPath = Path.Combine(Directory.GetCurrentDirectory(), odsFileName);

            if (!File.Exists(odsPath))
            {
                Console.WriteLine($"File not found: {odsPath}");
                return;
            }

            // Load the ODS workbook
            Workbook workbook = new Workbook(odsPath);

            // Check if the workbook contains a VBA project (macro)
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Determine whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA Project Signed: " + isSigned);

                // If signed, verify whether the signature is valid
                if (isSigned)
                {
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("VBA Project Signature Valid: " + isValid);
                }
                else
                {
                    Console.WriteLine("VBA Project is not signed; signature validity check is not applicable.");
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project (no macros).");
            }
        }
    }
}