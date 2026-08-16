// Title: Digitally sign an Excel VBA project with a password‑protected PFX certificate using Aspose.Cells for .NET
// Description: Creates a macro‑enabled workbook, optionally adds a VBA module, loads a PFX certificate secured by a strong password, builds a DigitalSignature with comments and timestamp, signs the workbook's VbaProject, and saves the signed file as an XLSM workbook.
// Keywords: Aspose.Cells VBA signing | C# PFX certificate Excel | digital signature macro-enabled workbook | sign VBA project .NET | load password protected PFX C# | Excel macro integrity | Aspose.Cells digital signature example
// Common Searches: how to sign a VBA project in an XLSM file using Aspose.Cells | C# code to apply a digital signature to an Excel macro project | load a password protected PFX and sign a VBA project in .NET | sign macro‑enabled workbook with certificate Aspose.Cells example | Aspose.Cells sign VBA project programmatically
// Developer Intent: Add a digital signature to a VBA project in an XLSM workbook by loading a password‑protected PFX certificate with Aspose.Cells for .NET.
// Use Cases: Guarantee macro integrity and authenticity before distributing Excel files. | Automate compliance by embedding corporate certificates in generated reports. | Validate VBA code authenticity in CI/CD pipelines that produce macro‑enabled workbooks.
// AI Prompts: Generate C# code that loads a PFX file with a password and signs an Aspose.Cells VbaProject. | Explain error handling for missing certificate files or incorrect passwords when signing a VBA project. | Provide a step‑by‑step tutorial to create a macro‑enabled workbook, add a VBA module, and apply a digital signature using Aspose.Cells.

using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

// Creates a macro‑enabled workbook, optionally adds a VBA module, loads a PFX certificate secured by a strong password, builds a DigitalSignature with comments and timestamp, signs the workbook's VbaProject, and saves the signed file as an XLSM workbook.
public class VbaProjectSignDemo
{
    public static void Main()
    {
        try
        {
            Run();
            Console.WriteLine("VBA project signed and saved successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        // Create a new workbook (in-memory)
        Workbook workbook = new Workbook();

        // Ensure the workbook has a VBA project by saving as a macro-enabled file and reloading it
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);
        File.Delete(tempFile);

        // (Optional) Add a VBA module so the project contains code
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
        workbook.VbaProject.Modules[moduleIndex].Codes =
            "Sub Demo()\r\n" +
            "    MsgBox \"Hello from VBA!\"\r\n" +
            "End Sub";

        // Load the PFX certificate that contains the private key
        string certPath = "mycert.pfx";
        string certPassword = "StrongPassword";

        if (!File.Exists(certPath))
            throw new FileNotFoundException($"Certificate file not found: {certPath}");

        X509Certificate2 certificate = new X509Certificate2();
        certificate.Import(certPath, certPassword, X509KeyStorageFlags.DefaultKeySet);

        // Create a digital signature using the certificate
        DigitalSignature digitalSignature = new DigitalSignature(
            certificate,               // certificate with private key
            "VBA Project Signature",   // comments / purpose
            DateTime.Now);             // signing time

        // Sign the VBA project
        workbook.VbaProject.Sign(digitalSignature);

        // Save the signed workbook as a macro-enabled file
        string outputPath = "SignedVbaProject.xlsm";
        workbook.Save(outputPath, SaveFormat.Xlsm);
    }
}
