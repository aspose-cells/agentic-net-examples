using System;
using System.Collections;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsExamples
{
    public class VbaSignatureVerificationDemo
    {
        public static void Run()
        {
            // Path to the macro-enabled workbook (xlsm) that contains the VBA project
            string workbookPath = "SignedWorkbook.xlsm";

            // Path to the TSV file that holds expected certificate data (Base64 encoded)
            string tsvPath = "CertificateData.tsv";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project is signed
            Console.WriteLine("VBA Project Signed: " + vbaProject.IsSigned);

            // If signed, verify the signature validity
            if (vbaProject.IsSigned)
            {
                Console.WriteLine("VBA Project Signature Valid: " + vbaProject.IsValidSigned);
            }
            else
            {
                Console.WriteLine("VBA Project is not signed; cannot verify signature.");
                return;
            }

            // Read expected certificate raw data from the TSV file (first column, first row)
            if (!File.Exists(tsvPath))
            {
                Console.WriteLine("TSV file not found: " + tsvPath);
                return;
            }

            string[] tsvLines = File.ReadAllLines(tsvPath);
            if (tsvLines.Length == 0)
            {
                Console.WriteLine("TSV file is empty.");
                return;
            }

            // Assume the first column contains Base64-encoded certificate bytes
            string base64Cert = tsvLines[0].Split('\t')[0];
            byte[] expectedCertData = Convert.FromBase64String(base64Cert);

            // Get the actual certificate raw data from the VBA project
            byte[] actualCertData = vbaProject.CertRawData;

            // Compare lengths and content as a simple validation step
            bool certMatches = actualCertData != null &&
                               expectedCertData.Length == actualCertData.Length &&
                               StructuralComparisons.StructuralEqualityComparer.Equals(expectedCertData, actualCertData);

            Console.WriteLine("Certificate data matches TSV entry: " + certMatches);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureVerificationDemo.Run();
        }
    }
}