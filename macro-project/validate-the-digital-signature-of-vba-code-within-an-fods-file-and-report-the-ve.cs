using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main(string[] args)
        {
            ValidateVbaSignatureInFods.Run();
        }
    }

    public class ValidateVbaSignatureInFods
    {
        public static void Run()
        {
            // Load the FODS (Flat OpenDocument Spreadsheet) file
            Workbook workbook = new Workbook("sample.fods");

            // Get the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Verify the validity of the signature
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}