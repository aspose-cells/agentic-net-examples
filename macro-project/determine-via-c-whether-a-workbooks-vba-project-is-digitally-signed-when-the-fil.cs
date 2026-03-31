using System;
using Aspose.Cells;

class CheckVbaSignatureInPdf
{
    static void Main()
    {
        // Path to the macro-enabled Excel file
        string excelPath = "sample.xlsm";

        // Load the workbook
        Workbook workbook = new Workbook(excelPath);

        // Determine if the VBA project exists and is digitally signed
        bool isVbaSigned = workbook.VbaProject != null && workbook.VbaProject.IsSigned;
        Console.WriteLine("VBA project signed: " + isVbaSigned);

        // If signed, optionally display whether the signature is valid
        if (isVbaSigned)
        {
            Console.WriteLine("Signature valid: " + workbook.VbaProject.IsValidSigned);
        }

        // Save the workbook as PDF (the PDF itself does not contain VBA, but this demonstrates the workflow)
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}