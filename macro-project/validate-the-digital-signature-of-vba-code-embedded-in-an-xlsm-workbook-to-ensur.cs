using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class ValidateVbaSignatureDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Load a macro-enabled workbook that contains a VBA project
            Workbook workbook = new Workbook("SignedVbaWorkbook.xlsm");

            // Get the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Validate the digital signature of the VBA project
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            // Save a copy of the workbook to preserve the signature (optional)
            workbook.Save("CopySignedWorkbook.xlsm", SaveFormat.Xlsm);
        }
    }
}