using System;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main()
        {
            // Load the Excel workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Access the VBA project associated with the workbook
            var vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            // If signed, optionally check if the signature is valid
            if (isSigned)
            {
                Console.WriteLine("Signature valid: " + vbaProject.IsValidSigned);
            }
        }
    }
}