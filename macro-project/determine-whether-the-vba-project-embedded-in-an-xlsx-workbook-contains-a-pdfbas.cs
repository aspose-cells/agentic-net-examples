using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the XLSX workbook to be examined
        string workbookPath = "sample.xlsx";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(workbookPath);

        // Verify that the workbook actually contains a VBA project (macro)
        if (workbook.HasMacro && workbook.VbaProject != null)
        {
            Console.WriteLine("Workbook contains a VBA project.");

            // Determine whether the VBA project is signed
            bool isSigned = workbook.VbaProject.IsSigned;
            Console.WriteLine("VBA project is signed: " + isSigned);

            if (isSigned)
            {
                // Check if the embedded signature is valid
                bool isValid = workbook.VbaProject.IsValidSigned;
                Console.WriteLine("VBA project signature is valid: " + isValid);

                // Retrieve raw certificate data (optional, for further inspection)
                byte[] certData = workbook.VbaProject.CertRawData;
                Console.WriteLine("Certificate raw data length: " + (certData?.Length ?? 0));
            }
        }
        else
        {
            Console.WriteLine("Workbook does not contain a VBA project.");
        }
    }
}