using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load a macro‑enabled workbook that may contain a VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Determine whether the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;
        Console.WriteLine("VBA project signed: " + isSigned);

        // If it is signed, also display whether the signature is valid
        if (isSigned)
        {
            Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
        }

        // Save the workbook in XPS format (OXPS is the Open XML Paper Specification,
        // which Aspose.Cells represents via the Xps save option)
        workbook.Save("sample_output.oxps", SaveFormat.Xps);
    }
}