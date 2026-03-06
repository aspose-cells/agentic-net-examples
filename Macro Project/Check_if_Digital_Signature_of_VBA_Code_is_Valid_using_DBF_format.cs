using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the DBF file (or any Excel file that may contain a VBA project)
            string inputPath = "sample.dbf";

            // Load the workbook using Aspose.Cells (DBF format is supported)
            Workbook workbook = new Workbook(inputPath);

            // Get the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Verify that a VBA project exists and whether it is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                // Check if the signature of the VBA project is valid
                Console.WriteLine("Signature valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or does not exist.");
            }
        }
    }
}