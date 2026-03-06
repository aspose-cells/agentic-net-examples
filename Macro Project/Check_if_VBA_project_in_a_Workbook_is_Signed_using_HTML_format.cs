using System;
using System.IO;
using Aspose.Cells;

class CheckVbaSignatureHtml
{
    static void Main()
    {
        // Load the workbook that may contain a VBA project
        Workbook workbook = new Workbook("sample.xlsm");

        // Check if the VBA project is signed
        bool isSigned = workbook.VbaProject.IsSigned;
        bool isValidSigned = false;

        if (isSigned)
        {
            // If signed, also check whether the signature is valid
            isValidSigned = workbook.VbaProject.IsValidSigned;
        }

        // Build a simple HTML report
        string html = "<!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>VBA Signature Status</title></head><body>";
        html += "<h1>VBA Project Signature Status</h1>";
        html += "<p>Signed: " + (isSigned ? "Yes" : "No") + "</p>";

        if (isSigned)
        {
            html += "<p>Signature Valid: " + (isValidSigned ? "Yes" : "No") + "</p>";
        }

        html += "</body></html>";

        // Save the HTML report to a file
        File.WriteAllText("VbaSignatureStatus.html", html);
    }
}