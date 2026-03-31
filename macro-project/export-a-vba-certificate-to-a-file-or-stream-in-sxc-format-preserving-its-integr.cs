using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ExportVbaCertificate
{
    public static void Run()
    {
        // Load a workbook that contains a signed VBA project
        string sourcePath = "SignedWorkbook.xlsm";
        Workbook workbook = new Workbook(sourcePath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // If the VBA project is signed, export its certificate raw data
        if (vbaProject.IsSigned && vbaProject.CertRawData != null && vbaProject.CertRawData.Length > 0)
        {
            // Save the certificate to a .cer file
            File.WriteAllBytes("VbaCertificate.cer", vbaProject.CertRawData);
        }

        // Save the workbook in StarOffice Calc (SXC) format, preserving the VBA signature
        workbook.Save("SignedWorkbook.sxc", SaveFormat.Sxc);
    }
}

class Program
{
    static void Main(string[] args)
    {
        ExportVbaCertificate.Run();
    }
}