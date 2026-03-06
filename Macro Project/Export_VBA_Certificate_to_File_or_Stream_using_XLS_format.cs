using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

public class ExportVbaCertificateDemo
{
    public static void Run()
    {
        // Load a workbook that contains a signed VBA project
        string sourcePath = "SignedWorkbook.xlsm";
        using (Workbook workbook = new Workbook(sourcePath))
        {
            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Export the certificate if the project is signed
            if (vbaProject != null && vbaProject.IsSigned)
            {
                byte[] certData = vbaProject.CertRawData;

                if (certData != null && certData.Length > 0)
                {
                    // Save certificate directly to a file
                    File.WriteAllBytes("VbaCertificate.cer", certData);

                    // Also demonstrate saving the certificate to a memory stream
                    using (MemoryStream certStream = new MemoryStream())
                    {
                        certStream.Write(certData, 0, certData.Length);
                        certStream.Position = 0; // reset for reading if needed

                        // Example: copy the stream to another file
                        using (FileStream fs = new FileStream("VbaCertificateFromStream.cer", FileMode.Create, FileAccess.Write))
                        {
                            certStream.CopyTo(fs);
                        }
                    }
                }
            }

            // Save the workbook as an Excel 97-2003 XLS file into a memory stream
            using (MemoryStream xlsStream = new MemoryStream())
            {
                var saveOptions = new XlsSaveOptions(); // default saves as XLS (Excel 97-2003)
                workbook.Save(xlsStream, saveOptions);
                xlsStream.Position = 0;

                // Write the XLS stream to a physical file
                using (FileStream file = new FileStream("WorkbookExported.xls", FileMode.Create, FileAccess.Write))
                {
                    xlsStream.CopyTo(file);
                }
            }
        }
    }
}

public class Program
{
    public static void Main()
    {
        ExportVbaCertificateDemo.Run();
    }
}