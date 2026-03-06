using System;
using System.IO;
using Aspose.Cells;

class CheckVbaProjectSignature
{
    static void Main()
    {
        // Path to the workbook (adjust as needed)
        string filePath = Path.Combine(Directory.GetCurrentDirectory(), "template.xltx");

        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Load the workbook
        Workbook workbook = new Workbook(filePath);

        // Determine whether the workbook contains a VBA project
        if (workbook.HasMacro)
        {
            // Check if the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project is signed: " + isSigned);

            // If signed, optionally verify the signature validity
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