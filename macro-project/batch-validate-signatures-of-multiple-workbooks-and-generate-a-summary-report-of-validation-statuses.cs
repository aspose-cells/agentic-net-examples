using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace SignatureBatchValidator
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the workbook files to validate.
            // In a real scenario you might read these from a directory or input arguments.
            string[] workbookPaths = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // List to hold summary lines.
            List<string> reportLines = new List<string>();
            reportLines.Add("FileName,IsSigned,SignatureCount,ValidSignatureCount");

            foreach (string path in workbookPaths)
            {
                // Load the workbook using the provided constructor (load rule).
                Workbook wb = new Workbook(path);

                // Determine if the workbook is digitally signed.
                bool isSigned = wb.IsDigitallySigned;

                int signatureCount = 0;
                int validSignatureCount = 0;

                if (isSigned)
                {
                    // Retrieve the digital signature collection (method rule).
                    DigitalSignatureCollection signatures = wb.GetDigitalSignature();

                    if (signatures != null)
                    {
                        foreach (DigitalSignature sig in signatures)
                        {
                            signatureCount++;
                            // Check each signature's validity (property rule).
                            if (sig.IsValid)
                            {
                                validSignatureCount++;
                            }
                        }
                    }
                }

                // Build a CSV line for the current workbook.
                string fileName = Path.GetFileName(path);
                string line = $"{fileName},{isSigned},{signatureCount},{validSignatureCount}";
                reportLines.Add(line);
            }

            // Write the summary report to a text file.
            string reportPath = "SignatureSummary.txt";
            File.WriteAllLines(reportPath, reportLines);

            Console.WriteLine($"Signature validation summary written to: {reportPath}");
        }
    }
}