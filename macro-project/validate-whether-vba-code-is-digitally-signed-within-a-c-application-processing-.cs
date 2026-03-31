using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled) to be examined
            string filePath = "sample.xlsm";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(filePath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Prepare output in TAB-separated format
            // Columns: FilePath    IsVbaSigned    IsVbaSignatureValid    WorkbookDigitallySigned
            string result = string.Join("\t",
                filePath,
                vbaProject.IsSigned.ToString(),
                vbaProject.IsValidSigned.ToString(),
                workbook.IsDigitallySigned.ToString());

            Console.WriteLine(result);

            // Optional: if VBA project is signed, display certificate raw data length
            if (vbaProject.IsSigned)
            {
                byte[] certData = vbaProject.CertRawData;
                Console.WriteLine("Certificate raw data length: " + (certData?.Length ?? 0));
            }
        }
    }
}