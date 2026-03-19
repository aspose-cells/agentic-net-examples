using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLTM workbook (macro-enabled template)
            string workbookPath = "sample.xltm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project exists
            if (vbaProject == null)
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
                return;
            }

            // Determine whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + isSigned);

            // If signed, check whether the signature is valid
            if (isSigned)
            {
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA Project Signature Valid: " + isValid);
            }
        }
    }
}