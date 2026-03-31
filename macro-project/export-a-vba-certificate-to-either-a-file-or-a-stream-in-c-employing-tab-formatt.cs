using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    static void Main()
    {
        // Path to a macro‑enabled workbook that contains a signed VBA project
        string signedWorkbookPath = "SignedWorkbook.xlsm";

        // Load the workbook (load rule)
        Workbook workbook = new Workbook(signedWorkbookPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Ensure the VBA project is signed
        if (!vbaProject.IsSigned)
        {
            Console.WriteLine("VBA project is not signed. No certificate to export.");
            return;
        }

        // Get the raw certificate data (binary)
        byte[] certData = vbaProject.CertRawData;

        if (certData == null || certData.Length == 0)
        {
            Console.WriteLine("Certificate data is empty.");
            return;
        }

        // -------------------------------------------------
        // Export certificate to a physical file
        // -------------------------------------------------
        string certFilePath = "VbaCertificate.cer";
        File.WriteAllBytes(certFilePath, certData);
        Console.WriteLine("Certificate saved to file: " + certFilePath);

        // -------------------------------------------------
        // Export certificate to a memory stream
        // -------------------------------------------------
        using (MemoryStream certStream = new MemoryStream())
        {
            // Write raw bytes to the stream
            certStream.Write(certData, 0, certData.Length);
            certStream.Position = 0; // reset for reading

            // Example: output the certificate as a TAB‑separated hex string
            Console.WriteLine("Certificate data (hex) with TAB separation:");
            byte[] buffer = certStream.ToArray();
            for (int i = 0; i < buffer.Length; i++)
            {
                Console.Write(buffer[i].ToString("X2"));
                if (i < buffer.Length - 1)
                    Console.Write("\t"); // TAB separator
            }
            Console.WriteLine();

            // The stream can now be returned or used elsewhere
        }
    }
}