using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load a macro-enabled workbook that may contain a VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Determine whether the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;
        Console.WriteLine("VBA project signed: " + isSigned);

        // If signed, optionally display whether the signature is valid
        if (isSigned)
        {
            Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
        }

        // Save the workbook in SXC format (preserves VBA project information)
        workbook.Save("temp.sxc", SaveFormat.SXC);

        // Reload the workbook from the SXC file to verify the signature status persists
        Workbook reloaded = new Workbook("temp.sxc");
        Console.WriteLine("After SXC reload - VBA project signed: " + reloaded.VbaProject.IsSigned);
    }
}