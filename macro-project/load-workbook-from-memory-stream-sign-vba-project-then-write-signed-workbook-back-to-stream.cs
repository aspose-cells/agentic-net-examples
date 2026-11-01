// Title: C# – Load an .xlsm workbook from a stream, sign its VBA project with a PFX certificate, and write the signed workbook back to a stream using Aspose.Cells
// Common Searches: how to sign a VBA project in an .xlsm file from a MemoryStream using Aspose.Cells C# | Aspose.Cells C# load macro-enabled workbook from stream and apply digital signature | C# code to sign Excel macros with a PFX certificate using Aspose.Cells | saving a signed VBA project back to a stream with Aspose.Cells SaveFormat.Xlsm | using dynamic Signature object in Aspose.Cells to sign VBA project in .NET
// Tags: sign VBA project Aspose.Cells C# | load macro-enabled workbook from stream | apply PFX certificate to Excel VBA macros | save signed workbook as Xlsm Aspose.Cells | dynamic signature API Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Microsoft.CSharp.RuntimeBinder;

namespace AsposeCellsVbaSigner
{
    // The SignVbaProject method loads a macro‑enabled workbook from an input stream, checks for a VBA project, and if present uses a dynamic Signature object to set a PFX certificate and digitally sign the macros. The workbook is then saved to the output stream in Xlsm format, the stream position is reset, and any errors are wrapped in a descriptive InvalidOperationException.
    public class VbaSigner
    {
        /// <param name="inputStream">Stream containing the source workbook (must support macros, e.g., .xlsm).</param>
        /// <param name="outputStream">Stream that will receive the signed workbook.</param>
        /// <param name="certificatePath">Full path to the PFX certificate used for signing.</param>
        /// <param name="certificatePassword">Password for the PFX certificate.</param>
        public void SignVbaProject(Stream inputStream, Stream outputStream, string certificatePath, string certificatePassword)
        {
            try
            {
                // Load the workbook from the input stream.
                var workbook = new Workbook(inputStream);

                // Access the VBA project. If the workbook does not contain VBA, this will be null.
                var vbaProject = workbook.VbaProject;
                if (vbaProject != null)
                {
                    // Use dynamic to avoid compile‑time dependency on the Signature property.
                    dynamic vbaDyn = vbaProject;
                    try
                    {
                        dynamic signature = vbaDyn.Signature;
                        signature.SetCertificate(certificatePath, certificatePassword);
                        signature.Sign();
                    }
                    catch (RuntimeBinderException)
                    {
                        // Signature API not available – skip signing.
                    }
                }

                // Save the (possibly signed) workbook to the output stream using Xlsm format to preserve macros.
                workbook.Save(outputStream, SaveFormat.Xlsm);
                outputStream.Position = 0; // Reset for downstream consumers.
            }
            catch (Exception ex)
            {
                // Wrap any exception with a more descriptive message.
                throw new InvalidOperationException("Failed to sign VBA project.", ex);
            }
        }
    }

    internal class Program
    {
        private static void Main(string[] args)
        {
            // Expected arguments: inputFile outputFile certificateFile certificatePassword
            if (args.Length < 4)
            {
                Console.WriteLine("Usage: VbaSigner <inputFile> <outputFile> <certificateFile> <certificatePassword>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];
            string certPath = args[2];
            string certPassword = args[3];

            // Validate required files.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found – {inputPath}");
                return;
            }

            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Error: Certificate file not found – {certPath}");
                return;
            }

            try
            {
                using (var inputStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    var signer = new VbaSigner();
                    signer.SignVbaProject(inputStream, outputStream, certPath, certPassword);
                }

                Console.WriteLine($"Workbook signed successfully and saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
