using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureDemo
{
    public class VbaSignatureValidator
    {
        public static void Run()
        {
            string workbookPath = "SignedVbaProject.xlsm";

            Workbook workbook = new Workbook(workbookPath);

            VbaProject vbaProject = workbook.VbaProject;

            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA project is signed.");

                bool isSignatureValid = vbaProject.IsValidSigned;
                Console.WriteLine("Signature valid: " + isSignatureValid);

                byte[] certData = vbaProject.CertRawData;
                if (certData != null && certData.Length > 0)
                {
                    string certPath = "VbaCertificate.cer";
                    File.WriteAllBytes(certPath, certData);
                    Console.WriteLine($"Certificate raw data saved to '{certPath}'.");
                }
            }
            else
            {
                Console.WriteLine("VBA project is not signed.");
            }

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0;

                Workbook reloadedWorkbook = new Workbook(ms);
                VbaProject reloadedVba = reloadedWorkbook.VbaProject;

                Console.WriteLine("After reload - VBA signed: " + reloadedVba.IsSigned);
                Console.WriteLine("After reload - Signature valid: " + reloadedVba.IsValidSigned);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            VbaSignatureValidator.Run();
        }
    }
}