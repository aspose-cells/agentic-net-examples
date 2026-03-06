using System;
using Aspose.Cells;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main()
        {
            // Load a macro-enabled workbook (replace with your file path)
            string workbookPath = "sample.xlsm";
            Workbook workbook = new Workbook(workbookPath);

            // Verify that the workbook actually contains a VBA project
            if (workbook.HasMacro && workbook.VbaProject != null)
            {
                // Check whether the VBA project is signed
                bool isSigned = workbook.VbaProject.IsSigned;
                Console.WriteLine("VBA project signed: " + isSigned);

                // If signed, also display whether the signature is valid
                if (isSigned)
                {
                    Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
                }
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }
        }
    }
}