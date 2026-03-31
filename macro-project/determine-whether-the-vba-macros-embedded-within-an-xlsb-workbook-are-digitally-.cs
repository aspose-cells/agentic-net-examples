using System;
using Aspose.Cells;

class CheckVbaSignature
{
    static void Main()
    {
        // Path to the XLSB workbook to be examined
        string workbookPath = "sample.xlsb";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(workbookPath);

        // Determine if the workbook contains any VBA macros
        if (workbook.HasMacro)
        {
            Console.WriteLine("Workbook contains VBA macros.");

            // Access the VBA project associated with the workbook
            var vbaProject = workbook.VbaProject;

            // Check whether the VBA project is signed
            Console.WriteLine("VBA Project IsSigned: " + vbaProject.IsSigned);

            // If signed, verify whether the signature is valid
            Console.WriteLine("VBA Project IsValidSigned: " + vbaProject.IsValidSigned);
        }
        else
        {
            Console.WriteLine("Workbook does not contain VBA macros.");
        }
    }
}