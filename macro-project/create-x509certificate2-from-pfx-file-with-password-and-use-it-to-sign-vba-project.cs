using System;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

class SignVbaProject
{
    static void Main()
    {
        try
        {
            // Verify input workbook exists
            string inputFile = "input.xlsm";
            if (!File.Exists(inputFile))
                throw new FileNotFoundException($"Input file not found: {inputFile}");

            // Load workbook containing VBA project
            Workbook workbook = new Workbook(inputFile);

            // Access VBA project
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject != null)
            {
                // Verify certificate file exists
                string certPath = "mycertificate.pfx";
                if (!File.Exists(certPath))
                    throw new FileNotFoundException($"Certificate file not found: {certPath}");

                // Load X509 certificate (replace password with the correct one)
                string certPassword = "yourPassword";
                X509Certificate2 certificate;
                try
                {
                    certificate = new X509Certificate2(certPath, certPassword, X509KeyStorageFlags.MachineKeySet);
                }
                catch (CryptographicException ex)
                {
                    Console.WriteLine($"Failed to load certificate: {ex.Message}");
                    return;
                }

                // Create digital signature
                DigitalSignature signature = new DigitalSignature(
                    certificate,
                    "Signed by Aspose.Cells",
                    DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(signature);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }

            // Save signed workbook
            string outputFile = "signed_output.xlsm";
            workbook.Save(outputFile, SaveFormat.Xlsm);
            Console.WriteLine($"Signed workbook saved to: {outputFile}");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}