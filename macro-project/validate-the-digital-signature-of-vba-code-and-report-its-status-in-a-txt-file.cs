using System;
using System.IO;
using Aspose.Cells;

class ValidateVbaSignature
{
    static void Main()
    {
        // Load the workbook that may contain a signed VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Determine the signature status of the VBA project
        string result;
        if (workbook.VbaProject.IsSigned)
        {
            // VBA project is signed; check if the signature is valid
            result = $"VBA Project is signed. Signature valid: {workbook.VbaProject.IsValidSigned}";
        }
        else
        {
            // No signature present
            result = "VBA Project is not signed.";
        }

        // Write the status to a text file
        File.WriteAllText("VbaSignatureStatus.txt", result);

        // Optional: display the result in the console
        Console.WriteLine(result);
    }
}