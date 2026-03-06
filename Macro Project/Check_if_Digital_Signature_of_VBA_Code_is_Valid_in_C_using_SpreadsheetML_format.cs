using System;
using System.IO;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load a macro-enabled workbook that may contain a signed VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Determine whether the VBA project is signed
        Console.WriteLine("VBA Project Signed: " + workbook.VbaProject.IsSigned);

        // If it is signed, check whether the signature is valid
        if (workbook.VbaProject.IsSigned)
        {
            Console.WriteLine("VBA Signature Valid: " + workbook.VbaProject.IsValidSigned);
        }

        // Save the workbook to a memory stream to preserve the signature
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsm);

            // Reload the workbook from the stream to verify that the signature persists
            Workbook reloadedWorkbook = new Workbook(stream);
            Console.WriteLine("After reload - VBA Project Signed: " + reloadedWorkbook.VbaProject.IsSigned);
            Console.WriteLine("After reload - VBA Signature Valid: " + reloadedWorkbook.VbaProject.IsValidSigned);
        }
    }
}