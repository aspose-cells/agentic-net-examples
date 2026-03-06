using System;
using Aspose.Cells;

class VbaSignaturePdfReport
{
    static void Main()
    {
        // Load the Excel workbook that may contain a VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Check if the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;

        // If signed, also check whether the signature is valid
        bool isValidSigned = false;
        if (isSigned)
        {
            isValidSigned = workbook.VbaProject.IsValidSigned;
        }

        // Create a new workbook that will be used to generate the PDF report
        Workbook reportWorkbook = new Workbook();
        Worksheet sheet = reportWorkbook.Worksheets[0];

        // Write the results into cells
        sheet.Cells["A1"].PutValue("VBA Project Signed:");
        sheet.Cells["B1"].PutValue(isSigned);
        sheet.Cells["A2"].PutValue("Signature Valid:");
        sheet.Cells["B2"].PutValue(isValidSigned);

        // Save the report as a PDF file
        reportWorkbook.Save("VbaSignatureReport.pdf", SaveFormat.Pdf);
    }
}