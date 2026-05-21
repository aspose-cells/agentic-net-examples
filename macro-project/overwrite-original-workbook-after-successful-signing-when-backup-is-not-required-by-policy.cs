using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.DigitalSignatures;

namespace AsposeCellsExamples
{
    public class SignAndOverwriteWorkbook
    {
        /// <summary>
        /// Signs the workbook with a digital signature and overwrites the original file.
        /// No backup copy is created.
        /// </summary>
        /// <param name="workbookPath">Full path to the workbook to be signed.</param>
        /// <param name="certificatePath">Full path to the PFX certificate file.</param>
        /// <param name="certificatePassword">Password for the certificate.</param>
        public static void Run(string workbookPath, string certificatePath, string certificatePassword)
        {
            try
            {
                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                    throw new FileNotFoundException("Workbook file not found.", workbookPath);

                // Verify that the certificate file exists
                if (!File.Exists(certificatePath))
                    throw new FileNotFoundException("Certificate file not found.", certificatePath);

                // Load the existing workbook
                Workbook workbook = new Workbook(workbookPath);

                // Load the certificate data (PFX file)
                byte[] certData = File.ReadAllBytes(certificatePath);

                // Create a digital signature using the certificate data
                DigitalSignature signature = new DigitalSignature(certData, certificatePassword, "Aspose.Signature", DateTime.Now);

                // Create a collection and add the signature
                DigitalSignatureCollection signatures = new DigitalSignatureCollection();
                signatures.Add(signature);

                // Apply the digital signature to the workbook
                workbook.SetDigitalSignature(signatures);

                // Overwrite the original workbook file with the signed version
                workbook.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error signing workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            // Expected arguments: workbookPath, certificatePath, certificatePassword
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: AsposeCellsRunner <workbookPath> <certificatePath> <certificatePassword>");
                return;
            }

            string workbookPath = args[0];
            string certificatePath = args[1];
            string certificatePassword = args[2];

            SignAndOverwriteWorkbook.Run(workbookPath, certificatePath, certificatePassword);
        }
    }
}