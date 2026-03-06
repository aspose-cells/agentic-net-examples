using System;
using Aspose.Cells;

namespace AsposeCellsVbaCheck
{
    class Program
    {
        static void Main()
        {
            // Path to the XLSB workbook
            string filePath = "sample.xlsb";

            // Load the workbook (XLSB format)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project
            var vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }
        }
    }
}