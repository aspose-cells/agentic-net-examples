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
            // Load the macro-enabled template workbook (XLTM)
            Workbook workbook = new Workbook("template.xltm");

            // Get the VBA project from the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is signed and whether the signature is valid
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}