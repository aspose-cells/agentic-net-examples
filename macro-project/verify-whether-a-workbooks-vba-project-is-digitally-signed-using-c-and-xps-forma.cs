using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VerifyVbaSignatureXpsDemo
    {
        public static void Main()
        {
            // Load a macro-enabled workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Determine whether the workbook has a VBA project and if it is signed
            bool vbaSigned = workbook.VbaProject != null && workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA Project Signed: " + vbaSigned);

            // If signed, verify that the signature is valid
            if (vbaSigned)
            {
                Console.WriteLine("VBA Signature Valid: " + workbook.VbaProject.IsValidSigned);
            }

            // Save the workbook as XPS (visual representation of the spreadsheet)
            workbook.Save("output.xps", SaveFormat.Xps);
        }
    }
}