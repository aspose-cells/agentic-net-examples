using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    class Program
    {
        static void Main()
        {
            VerifyVbaSignatureDemo.Run();
        }
    }

    public class VerifyVbaSignatureDemo
    {
        public static void Run()
        {
            // Load the Excel workbook that may contain a VBA project
            Workbook workbook = new Workbook("sample.xlsm");

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Check whether the VBA project is digitally signed
            bool isSigned = vbaProject.IsSigned;
            Console.WriteLine("VBA project signed: " + isSigned);

            if (isSigned)
            {
                // If signed, verify whether the signature is valid
                bool isValid = vbaProject.IsValidSigned;
                Console.WriteLine("VBA signature valid: " + isValid);

                // Retrieve the raw certificate data for further inspection
                byte[] certData = vbaProject.CertRawData;
                Console.WriteLine("Certificate raw data length: " + (certData?.Length ?? 0));
            }
            else
            {
                Console.WriteLine("VBA project is not signed; no signature to validate.");
            }
        }
    }
}