using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureInDifDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Load a workbook from a DIF file
            // Replace "sample.dif" with the actual path to your DIF file
            Workbook workbook = new Workbook("sample.dif");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("The VBA project is signed.");
                Console.WriteLine("Is the signature valid? " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("The VBA project is not signed or does not exist.");
            }
        }
    }
}