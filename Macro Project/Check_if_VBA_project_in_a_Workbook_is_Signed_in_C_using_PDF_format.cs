using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main()
        {
            string inputPath = "SampleWithVba.xlsm";

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook = new Workbook(inputPath);

            bool isSigned = workbook.VbaProject.IsSigned;

            Console.WriteLine("VBA project signed: " + isSigned);
            if (isSigned)
            {
                Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
            }

            Worksheet resultSheet = workbook.Worksheets.Add("SignatureInfo");

            resultSheet.Cells["A1"].PutValue("VBA Project Signed:");
            resultSheet.Cells["B1"].PutValue(isSigned);
            resultSheet.Cells["A2"].PutValue("Signature Valid:");
            resultSheet.Cells["B2"].PutValue(isSigned ? workbook.VbaProject.IsValidSigned : false);

            string outputPdf = "VbaSignatureInfo.pdf";
            workbook.Save(outputPdf, SaveFormat.Pdf);

            Console.WriteLine("Result saved to PDF: " + outputPdf);
        }
    }
}