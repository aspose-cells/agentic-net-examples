using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file that contains the VBA project (e.g., .xlsm)
            string workbookPath = "SignedVbaWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Validate the digital signature of the VBA project
                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("Is VBA signature valid? " + isSignatureValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            // Optional: Reload the workbook to ensure the signature status persists
            using (var ms = new System.IO.MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0;
                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVba = reloadedWorkbook.VbaProject;

                Console.WriteLine("After reload - Is VBA signed? " + reloadedVba.IsSigned);
                Console.WriteLine("After reload - Is VBA signature valid? " + reloadedVba.IsValidSigned);
            }
        }
    }
}