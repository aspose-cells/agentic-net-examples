using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaCertificateExport
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains a signed VBA project
            string signedWorkbookPath = "SignedWithVba.xlsm";

            // Ensure the workbook exists; create a placeholder if it does not
            if (!File.Exists(signedWorkbookPath))
            {
                Workbook placeholder = new Workbook();
                placeholder.Save(signedWorkbookPath, SaveFormat.Xlsm);
                Console.WriteLine($"Placeholder workbook created at: {signedWorkbookPath}");
            }

            // Load the workbook
            Workbook workbook = new Workbook(signedWorkbookPath);

            // Access the VBA project (may be null if none exists)
            VbaProject vbaProject = workbook.VbaProject;

            // Check if the VBA project exists and is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                // Retrieve the raw certificate data
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // 1. Save the raw certificate bytes directly to a .cer file
                    string certFilePath = "VbaCertificate.cer";
                    File.WriteAllBytes(certFilePath, certData);
                    Console.WriteLine($"Certificate binary saved to: {certFilePath}");

                    // 2. Export the certificate data as a TAB‑separated text file
                    StringBuilder tabBuilder = new StringBuilder();
                    for (int i = 0; i < certData.Length; i++)
                    {
                        tabBuilder.Append(certData[i].ToString("X2"));
                        if (i < certData.Length - 1)
                            tabBuilder.Append('\t');
                    }

                    string tabFilePath = "VbaCertificate.tab";
                    File.WriteAllText(tabFilePath, tabBuilder.ToString());
                    Console.WriteLine($"Certificate exported as TAB text to: {tabFilePath}");

                    // 3. Write the same TAB text to a MemoryStream
                    using (MemoryStream memStream = new MemoryStream())
                    {
                        byte[] tabBytes = Encoding.UTF8.GetBytes(tabBuilder.ToString());
                        memStream.Write(tabBytes, 0, tabBytes.Length);
                        memStream.Position = 0;

                        string streamOutputPath = "VbaCertificateFromStream.tab";
                        using (FileStream fileStream = new FileStream(streamOutputPath, FileMode.Create, FileAccess.Write))
                        {
                            memStream.CopyTo(fileStream);
                        }
                        Console.WriteLine($"Certificate TAB data written from stream to: {streamOutputPath}");
                    }
                }
                else
                {
                    Console.WriteLine("Certificate data is empty.");
                }
            }
            else
            {
                Console.WriteLine("The VBA project is not present or not signed; no certificate available.");
            }

            // Optional: Save the workbook after any modifications
            workbook.Save("ProcessedWorkbook.xlsm", SaveFormat.Xlsm);
        }
    }
}