using System;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureCheck
{
    public class VbaSignatureChecker
    {
        // Determines whether the VBA project in the specified workbook is digitally signed.
        public static bool IsVbaSigned(string filePath)
        {
            // Load the workbook from the given file path.
            Workbook workbook = new Workbook(filePath);

            // Verify that a VBA project exists and check its signing status.
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
                return true;
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
                return false;
            }
        }

        // Example entry point.
        public static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: VbaSignatureChecker <path-to-workbook>");
                return;
            }

            string filePath = args[0];
            IsVbaSigned(filePath);
        }
    }
}