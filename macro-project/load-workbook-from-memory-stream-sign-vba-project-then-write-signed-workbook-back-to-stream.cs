// Title: C# – Sign VBA Project in an XLSM Workbook from a MemoryStream with Aspose.Cells
// Description: Shows how to load a macro‑enabled XLSM workbook from a seekable stream, import a .pfx certificate, create a DigitalSignature, sign the VBA project when it is unsigned, and write the signed workbook back to a MemoryStream using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | sign VBA project | macro enabled workbook | XLSM | MemoryStream | digital signature | PFX certificate | VBA signing | load workbook from stream | save workbook to stream
// Common Searches: how to digitally sign a VBA project in an XLSM file using Aspose.Cells | load workbook from MemoryStream, sign macros, and save back to stream C# | apply a .pfx certificate to a VBA project with Aspose.Cells for .NET | sign macro‑enabled Excel file programmatically in C# | Aspose.Cells VBA project signing example
// Developer Intent: Load an XLSM workbook from a stream, apply a digital signature to its VBA project using a .pfx certificate, and obtain the signed workbook as a MemoryStream.
// Use Cases: Securely sign workbooks received via a web API before persisting them in a document store. | Enforce macro‑security compliance for batch‑generated reports on a server. | Create on‑the‑fly signed macro‑enabled Excel files for downstream applications that require trusted VBA code.
// AI Prompts: Write C# code that loads an XLSM workbook from a MemoryStream, signs its VBA project with a .pfx certificate using Aspose.Cells, and returns the signed workbook as a MemoryStream. | Explain how to check whether a VBA project is unsigned before applying a DigitalSignature with Aspose.Cells for .NET. | Provide best‑practice error handling for signing a VBA project from a stream and saving the result back to a stream.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

// Shows how to load a macro‑enabled XLSM workbook from a seekable stream, import a .pfx certificate, create a DigitalSignature, sign the VBA project when it is unsigned, and write the signed workbook back to a MemoryStream using Aspose.Cells for .NET.
public class VbaSigningHelper
{
    /// <param name="inputStream">Stream containing the original workbook (must support seeking).</param>
    /// <param name="certificatePath">Full path to the .pfx certificate file.</param>
    /// <param name="certificatePassword">Password for the .pfx certificate.</param>
    /// <returns>A MemoryStream containing the signed workbook.</returns>
    public static MemoryStream SignVbaProjectFromStream(Stream inputStream, string certificatePath, string certificatePassword)
    {
        try
        {
            // Ensure the input stream is positioned at the beginning
            if (inputStream.CanSeek)
                inputStream.Position = 0;

            // Load the workbook from the input stream
            Workbook workbook = new Workbook(inputStream);

            // Access the VBA project (may be null if the workbook has no macros)
            VbaProject vbaProject = workbook.VbaProject;
            if (vbaProject != null && !vbaProject.IsSigned)
            {
                // Verify certificate file exists
                if (!File.Exists(certificatePath))
                    throw new FileNotFoundException("Certificate file not found.", certificatePath);

                // Load the signing certificate without using the obsolete constructor
                X509Certificate2 certificate = new X509Certificate2();
                certificate.Import(certificatePath, certificatePassword, X509KeyStorageFlags.DefaultKeySet);

                // Create a digital signature instance
                DigitalSignature signature = new DigitalSignature(certificate, "Signed by Aspose.Cells", DateTime.Now);

                // Sign the VBA project
                vbaProject.Sign(signature);
            }

            // Prepare an output stream to hold the signed workbook
            MemoryStream outputStream = new MemoryStream();

            // Save the workbook in macro‑enabled format (XLSM) to the output stream
            workbook.Save(outputStream, SaveFormat.Xlsm);

            // Reset the position so the caller can read from the beginning
            outputStream.Position = 0;

            return outputStream;
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to preserve stack trace for callers
            throw new InvalidOperationException("Failed to sign VBA project.", ex);
        }
    }
}

public class Program
{
    /// <summary>
    /// Demonstrates signing a workbook's VBA project using a PFX certificate.
    /// </summary>
    public static void Main()
    {
        const string inputWorkbookPath = "SampleWithMacros.xlsm";
        const string certificatePath = "mycert.pfx";
        const string certificatePassword = "password";
        const string outputWorkbookPath = "SignedSample.xlsm";

        try
        {
            // Verify input workbook exists
            if (!File.Exists(inputWorkbookPath))
                throw new FileNotFoundException("Input workbook not found.", inputWorkbookPath);

            // Verify certificate file exists (additional check inside helper, but kept here for clarity)
            if (!File.Exists(certificatePath))
                throw new FileNotFoundException("Certificate file not found.", certificatePath);

            // Open the input workbook as a stream
            using (FileStream inputStream = new FileStream(inputWorkbookPath, FileMode.Open, FileAccess.Read))
            {
                // Sign the VBA project
                using (MemoryStream signedStream = VbaSigningHelper.SignVbaProjectFromStream(inputStream, certificatePath, certificatePassword))
                {
                    // Write the signed workbook to disk
                    using (FileStream outputStream = new FileStream(outputWorkbookPath, FileMode.Create, FileAccess.Write))
                    {
                        signedStream.CopyTo(outputStream);
                    }
                }
            }

            Console.WriteLine($"Workbook signed successfully. Output saved to '{outputWorkbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
