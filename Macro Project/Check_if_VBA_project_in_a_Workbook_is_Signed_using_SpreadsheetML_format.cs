using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProjectSignature
{
    static void Main()
    {
        // Load a macro-enabled workbook (XLSM) that may contain a VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Determine whether the workbook contains any macros/VBA code
        Console.WriteLine("Workbook contains macro: " + workbook.HasMacro);

        // Access the VBA project (property is always available; it may be empty)
        VbaProject vbaProject = workbook.VbaProject;

        // Check if the VBA project is signed
        bool isSigned = vbaProject.IsSigned;
        Console.WriteLine("VBA project is signed: " + isSigned);

        // If signed, optionally verify whether the signature is valid
        if (isSigned)
        {
            Console.WriteLine("VBA signature is valid: " + vbaProject.IsValidSigned);
        }
    }
}