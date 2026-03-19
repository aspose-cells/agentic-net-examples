using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;
using Aspose.Cells.Ods;

class ExportVbaCertificate
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Load a macro‑enabled workbook that contains a signed VBA project
        string inputPath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(inputPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the project is signed, export the certificate raw data to a .cer file
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            string certPath = "VbaCertificate.cer";
            File.WriteAllBytes(certPath, vbaProject.CertRawData);
        }

        // Save the workbook as ODS using OdsSaveOptions
        OdsSaveOptions odsOptions = new OdsSaveOptions
        {
            GeneratorType = OdsGeneratorType.LibreOffice // optional configuration
        };

        string odsPath = "SignedWorkbook.ods";
        workbook.Save(odsPath, odsOptions);
    }
}