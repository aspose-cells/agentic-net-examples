using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class CheckVbaSignatureDemo
    {
        public static void Run()
        {
            // Path to the Excel file that may contain a VBA project
            string excelPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Access the VBA project associated with the workbook
            var vbaProject = workbook.VbaProject;

            // Determine whether the VBA project is signed
            bool isSigned = vbaProject.IsSigned;

            // Prepare the result text
            string result = isSigned
                ? $"VBA project in '{excelPath}' is signed. Valid signature: {vbaProject.IsValidSigned}"
                : $"VBA project in '{excelPath}' is NOT signed.";

            // Output the result to the console
            Console.WriteLine(result);

            // Write the result to a TXT file
            string txtOutputPath = "VbaSignatureStatus.txt";
            File.WriteAllText(txtOutputPath, result);

            // Optionally, if the VBA project is signed, save the certificate raw data to a file
            if (isSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
            {
                string certPath = "VbaCertificate.cer";
                File.WriteAllBytes(certPath, vbaProject.CertRawData);
                Console.WriteLine($"Certificate raw data saved to '{certPath}'.");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CheckVbaSignatureDemo.Run();
        }
    }
}