using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLS workbook that may contain a VBA project
            string workbookPath = "sample.xls";

            // Load the workbook (XLS format)
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Optionally, verify if the signature is valid
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}