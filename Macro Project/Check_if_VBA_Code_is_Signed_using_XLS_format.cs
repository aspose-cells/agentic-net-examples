using System;
using Aspose.Cells;

class CheckVbaSigned
{
    static void Main()
    {
        // Load the Excel file in XLS format (Excel 97-2003)
        Workbook workbook = new Workbook("sample.xls");

        // Verify that the workbook contains a VBA project (macro)
        if (workbook.HasMacro)
        {
            // Determine whether the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project is signed: " + isSigned);

            // If signed, also report whether the signature is valid
            if (isSigned)
            {
                Console.WriteLine("Signature is valid: " + workbook.VbaProject.IsValidSigned);
            }
        }
        else
        {
            Console.WriteLine("The workbook does not contain a VBA project.");
        }
    }
}