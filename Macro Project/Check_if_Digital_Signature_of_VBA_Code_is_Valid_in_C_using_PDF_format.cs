using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    public class VbaSignatureValidator
    {
        public static void Run()
        {
            // Path to the Excel file that contains a VBA project (must be .xlsm)
            string excelPath = "SignedWorkbook.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                // Verify whether the signature is valid
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature is valid: " + isValid);
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            // Optionally, save the workbook as PDF (signature information is not transferred to PDF)
            string pdfPath = "SignedWorkbook.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine("Workbook saved as PDF to: " + pdfPath);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureValidator.Run();
        }
    }
}