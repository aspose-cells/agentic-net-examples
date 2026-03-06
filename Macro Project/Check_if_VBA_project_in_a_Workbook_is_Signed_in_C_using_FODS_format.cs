using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignedInFods
    {
        public static void Main()
        {
            // Load the workbook from a FODS (Flat OpenDocument Spreadsheet) file
            Workbook workbook = new Workbook("sample.fods");

            // Determine if the workbook contains a VBA project and whether it is signed
            if (workbook.HasMacro && workbook.VbaProject != null && workbook.VbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");
                Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA project is not signed or the workbook has no macros.");
            }
        }
    }
}