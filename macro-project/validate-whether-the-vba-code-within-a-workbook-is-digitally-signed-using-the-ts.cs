using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (must be macro-enabled, e.g., .xlsm)
            string workbookPath = "SignedMacroWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Verify that the workbook contains VBA/macros
            if (!workbook.HasMacro)
            {
                Console.WriteLine("The workbook does not contain any VBA project.");
                return;
            }

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Validate the signature
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("Is the VBA signature valid? " + isValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            // Optional: Save a copy to demonstrate that the signature status persists
            string copyPath = "CopyOfWorkbook.xlsm";
            workbook.Save(copyPath, SaveFormat.Xlsm);
            Console.WriteLine("Workbook saved to: " + copyPath);
        }
    }
}