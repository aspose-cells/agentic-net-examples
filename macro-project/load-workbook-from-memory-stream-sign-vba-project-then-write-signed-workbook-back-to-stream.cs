using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

public class VbaSigner
{
    /// <summary>
    /// Loads a workbook from <paramref name="inputStream"/>, signs its VBA project with the
    /// certificate located at <paramref name="certificatePath"/>, and writes the signed workbook
    /// to <paramref name="outputStream"/> in macro‑enabled XLSM format.
    /// </summary>
    public static void SignVbaProject(
        Stream inputStream,
        Stream outputStream,
        string certificatePath,
        string certificatePassword)
    {
        try
        {
            // Ensure the input stream is positioned at the beginning
            if (inputStream.CanSeek)
                inputStream.Position = 0;

            // Load the workbook from the provided stream
            var workbook = new Workbook(inputStream);

            // Access the VBA project (may be null if the workbook has no macros)
            var vbaProject = workbook.VbaProject;
            if (vbaProject != null)
            {
                // Verify certificate file exists
                if (!File.Exists(certificatePath))
                {
                    Console.WriteLine($"Certificate file not found: {certificatePath}");
                    return;
                }

                // Load the signing certificate (must contain a private key)
                var cert = new X509Certificate2(certificatePath, certificatePassword);

                // Create a DigitalSignature instance
                var signature = new DigitalSignature(cert, "Signed by Aspose.Cells", DateTime.UtcNow);

                // Sign the VBA project
                vbaProject.Sign(signature);
            }
            else
            {
                Console.WriteLine("The workbook does not contain a VBA project.");
            }

            // Save the signed workbook to the output stream in XLSM format
            workbook.Save(outputStream, SaveFormat.Xlsm);

            // Reset the output stream position for callers that will read from it
            if (outputStream.CanSeek)
                outputStream.Position = 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during VBA signing: {ex.Message}");
        }
    }
}

// Example usage
public class Program
{
    public static void Main()
    {
        try
        {
            const string inputPath = "UnsignedWorkbook.xlsm";
            const string outputPath = "SignedWorkbook.xlsm";
            const string certPath = "mycert.pfx";
            const string certPassword = "password";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Verify certificate exists
            if (!File.Exists(certPath))
            {
                Console.WriteLine($"Certificate file not found: {certPath}");
                return;
            }

            // Load input workbook into a memory stream
            using (var fileIn = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            using (var inputMs = new MemoryStream())
            {
                fileIn.CopyTo(inputMs);

                // Prepare output stream
                using (var outputMs = new MemoryStream())
                {
                    // Sign the VBA project
                    VbaSigner.SignVbaProject(inputMs, outputMs, certPath, certPassword);

                    // Write the signed workbook to a file for verification
                    File.WriteAllBytes(outputPath, outputMs.ToArray());

                    // Optional: verify signature status
                    var verifyWb = new Workbook(new MemoryStream(outputMs.ToArray()));
                    Console.WriteLine("Is Signed: " + verifyWb.VbaProject.IsSigned);
                    Console.WriteLine("Is Valid Signed: " + verifyWb.VbaProject.IsValidSigned);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}