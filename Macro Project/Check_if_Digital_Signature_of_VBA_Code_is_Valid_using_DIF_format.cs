using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            CheckVbaSignatureDemo.Run();
        }
    }

    public class CheckVbaSignatureDemo
    {
        public static void Run()
        {
            // Load a workbook that contains a VBA project (e.g., an .xlsm file)
            Workbook workbook = new Workbook("sample.xlsm");

            // Determine whether the VBA project is signed
            if (workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                // Verify if the VBA project's digital signature is valid
                Console.WriteLine("VBA signature valid: " + workbook.VbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}