using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSB workbook that may contain a signed VBA project
            string workbookPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SignedVbaWorkbook.xlsb");

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(workbookPath);

            // Get the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Check whether the digital signature of the VBA project is valid
                Console.WriteLine("Signature is valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}