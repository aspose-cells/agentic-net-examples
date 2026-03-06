using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureInOdsDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Load the ODS workbook (replace with actual file path)
            Workbook workbook = new Workbook("sample.ods");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // If signed, check whether the signature is valid
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValid);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed; no signature to validate.");
            }

            // Optionally, save the workbook to preserve any changes (not required for validation)
            // workbook.Save("output.ods", SaveFormat.Ods);
        }
    }
}