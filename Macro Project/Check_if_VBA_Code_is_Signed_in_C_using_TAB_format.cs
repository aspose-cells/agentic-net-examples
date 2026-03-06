using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

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
            // Load the workbook (macro-enabled file)
            Workbook workbook = new Workbook("sample.xlsm");

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                // If signed, also check whether the signature is valid
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}