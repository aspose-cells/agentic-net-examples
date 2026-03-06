using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignature
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the macro-enabled workbook (adjust as needed)
            string workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample_with_macro.xlsm");

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            Workbook workbook = new Workbook(workbookPath);

            // Verify that the workbook actually contains a VBA project
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Check whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If signed, also check whether the signature is valid
                if (isSigned)
                {
                    bool isValid = workbook.VbaProject.IsValidSigned;
                    Console.WriteLine("Signature valid: " + isValid);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }
}