// Title: Sign VBA Project in an XLSM from a MemoryStream and Return a Signed Stream – Aspose.Cells C#
// Description: Load an XLSM workbook from a MemoryStream, detect its VBA project, apply a digital signature using a .pfx certificate, and save the signed workbook back to a MemoryStream with Aspose.Cells for .NET.
// Keywords: Aspose.Cells VBA signing | C# sign Excel macro | digital signature XLSM | MemoryStream workbook | save signed workbook to stream | Xlsm digital certificate | Aspose.Cells C# example | VBA project sign programmatically
// Common Searches: how to digitally sign a VBA project in an XLSM using Aspose.Cells | load Excel workbook from MemoryStream, sign macros, and write back to stream C# | Aspose.Cells sign VBA macro with .pfx certificate | save signed Excel file to MemoryStream without temporary files | C# code to sign VBA project in memory
// Developer Intent: Apply a digital signature to the VBA project of an XLSM workbook loaded from a MemoryStream and obtain the signed workbook as a new MemoryStream.
// Use Cases: Web API that receives macro‑enabled Excel files, signs the VBA code in‑memory, and streams the signed file back to the client. | CI/CD step that batch‑processes XLSM files, adds a digital signature to each VBA project, and stores the results without creating intermediate files. | Desktop utility for bulk‑signing Excel workbooks using a user‑provided .pfx certificate while keeping all I/O operations in streams for performance.
// AI Prompts: Generate C# code that uses Aspose.Cells to sign a VBA project from a MemoryStream and returns the signed workbook as a MemoryStream. | Explain how to handle workbooks that lack a VBA project when applying a digital signature with Aspose.Cells. | Show how to extend the signing method to accept multiple certificates and customize the signature comment.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

namespace VbaSigningDemo
{
    // Load an XLSM workbook from a MemoryStream, detect its VBA project, apply a digital signature using a .pfx certificate, and save the signed workbook back to a MemoryStream with Aspose.Cells for .NET.
    public static class VbaSigner
    {
        /// <param name="inputStream">MemoryStream containing the original workbook.</param>
        /// <param name="certificatePath">Full path to the .pfx certificate file.</param>
        /// <param name="certificatePassword">Password for the certificate.</param>
        /// <returns>MemoryStream with the signed workbook (Xlsm format).</returns>
        public static MemoryStream SignVbaProject(MemoryStream inputStream, string certificatePath, string certificatePassword)
        {
            try
            {
                // Ensure the stream is at the beginning.
                inputStream.Position = 0;

                // Load workbook from the stream.
                Workbook workbook = new Workbook(inputStream);

                // Access VBA project (may be null if no macros).
                VbaProject vbaProject = workbook.VbaProject;

                if (vbaProject != null)
                {
                    if (!File.Exists(certificatePath))
                        throw new FileNotFoundException("Certificate file not found.", certificatePath);

                    // Load the signing certificate.
                    X509Certificate2 certificate = new X509Certificate2(certificatePath, certificatePassword);

                    // Create a digital signature.
                    DigitalSignature digitalSignature = new DigitalSignature(certificate, "Signed by Aspose.Cells", DateTime.Now);

                    // Sign the VBA project.
                    vbaProject.Sign(digitalSignature);
                }

                // Save signed workbook to a new memory stream.
                MemoryStream signedStream = new MemoryStream();
                workbook.Save(signedStream, SaveFormat.Xlsm);
                signedStream.Position = 0; // Reset for downstream reading.

                return signedStream;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error signing VBA project: {ex.Message}");
                throw;
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point. Expects: <inputXlsm> <certificatePath> <certificatePassword> <outputXlsm>
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: VbaSigningDemo <inputXlsm> <certificatePath> <certificatePassword> <outputXlsm>");
                return;
            }

            string inputPath = args[0];
            string certPath = args[1];
            string certPassword = args[2];
            string outputPath = args[3];

            try
            {
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException("Input workbook not found.", inputPath);
                if (!File.Exists(certPath))
                    throw new FileNotFoundException("Certificate file not found.", certPath);

                // Load input workbook into memory.
                using (FileStream fs = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
                using (MemoryStream inputMs = new MemoryStream())
                {
                    fs.CopyTo(inputMs);

                    // Sign the VBA project.
                    MemoryStream signedMs = VbaSigner.SignVbaProject(inputMs, certPath, certPassword);

                    // Write signed workbook to output file.
                    using (FileStream outFs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                    {
                        signedMs.CopyTo(outFs);
                    }
                }

                Console.WriteLine($"Workbook signed successfully. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Failed: {ex.Message}");
            }
        }
    }
}
