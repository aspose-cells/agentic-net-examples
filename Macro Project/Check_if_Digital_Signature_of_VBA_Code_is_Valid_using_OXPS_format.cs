using System;
using System.IO;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Load a workbook that contains a VBA project (e.g., an .xlsm file)
        Workbook workbook = new Workbook("SignedVbaWorkbook.xlsm");

        // Determine whether the VBA project is signed
        Console.WriteLine("VBA Project Signed: " + workbook.VbaProject.IsSigned);

        // If it is signed, verify the validity of the signature
        if (workbook.VbaProject.IsSigned)
        {
            Console.WriteLine("VBA Signature Valid: " + workbook.VbaProject.IsValidSigned);
        }

        // Optionally, check if the entire workbook is digitally signed
        Console.WriteLine("Workbook Digitally Signed: " + workbook.IsDigitallySigned);

        // Save the workbook to XPS format using a memory stream
        using (MemoryStream xpsStream = new MemoryStream())
        {
            workbook.Save(xpsStream, SaveFormat.Xps);
            Console.WriteLine("Workbook saved to XPS format, stream length: " + xpsStream.Length);
        }
    }
}