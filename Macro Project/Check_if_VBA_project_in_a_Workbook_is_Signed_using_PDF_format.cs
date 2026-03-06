using System;
using Aspose.Cells;

class CheckVbaSignedPdf
{
    static void Main()
    {
        // Load the Excel workbook that may contain a VBA project
        Workbook workbook = new Workbook("input.xlsm");

        // Check if the VBA project is signed and if the signature is valid
        bool isSigned = workbook.VbaProject.IsSigned;
        bool isValidSigned = isSigned && workbook.VbaProject.IsValidSigned;

        // Create a new workbook to hold the result report
        Workbook reportWorkbook = new Workbook();
        Worksheet sheet = reportWorkbook.Worksheets[0];

        // Write the check results into the report worksheet
        sheet.Cells["A1"].PutValue("VBA Project Signed:");
        sheet.Cells["B1"].PutValue(isSigned);
        sheet.Cells["A2"].PutValue("Signature Valid:");
        sheet.Cells["B2"].PutValue(isValidSigned);

        // Save the report as a PDF file
        reportWorkbook.Save("VbaSignatureReport.pdf", SaveFormat.Pdf);
    }
}