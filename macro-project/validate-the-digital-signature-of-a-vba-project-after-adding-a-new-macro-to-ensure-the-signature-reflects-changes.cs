using System;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.DigitalSignatures;

class ValidateVbaSignatureAfterMacro
{
    static void Main()
    {
        // Create a new workbook and save it as a macro‑enabled file to ensure a VBA project exists
        Workbook workbook = new Workbook();
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);

        // Reload the workbook so that the VBA project is initialized
        workbook = new Workbook(tempFile);

        // Add a new VBA module (class type) and insert macro code
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "NewModule");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];
        module.Codes = "Sub NewMacro()\r\n    MsgBox \"Hello from new macro!\"\r\nEnd Sub";

        // Load a signing certificate (replace with your own certificate path and password)
        X509Certificate2 certificate = new X509Certificate2("MyCertificate.pfx", "certPassword");

        // Create a digital signature for the VBA project
        DigitalSignature vbaSignature = new DigitalSignature(certificate, "VBA Project Signature", DateTime.Now);

        // Sign the VBA project with the created signature
        workbook.VbaProject.Sign(vbaSignature);

        // Save the workbook to a memory stream
        using (MemoryStream ms = new MemoryStream())
        {
            workbook.Save(ms, SaveFormat.Xlsm);

            // Reload the workbook from the stream to verify the signature
            Workbook reloadedWorkbook = new Workbook(ms);

            // Output signature validation results
            Console.WriteLine("VBA Project IsSigned: " + reloadedWorkbook.VbaProject.IsSigned);
            Console.WriteLine("VBA Project IsValidSigned: " + reloadedWorkbook.VbaProject.IsValidSigned);
        }

        // Clean up the temporary file
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }
    }
}