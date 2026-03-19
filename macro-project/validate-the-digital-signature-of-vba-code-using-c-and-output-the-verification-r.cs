using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class VbaSignatureValidationToHtml
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("signedWorkbook.xlsm");
            bool isSigned = workbook.VbaProject.IsSigned;
            bool isValidSigned = workbook.VbaProject.IsValidSigned;

            string html = "<!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>VBA Signature Validation</title></head><body>";
            html += "<h2>VBA Project Signature Validation</h2>";
            html += $"<p>Is Signed: {isSigned}</p>";
            html += $"<p>Is Signature Valid: {isValidSigned}</p>";
            html += "</body></html>";

            File.WriteAllText("VbaSignatureResult.html", html);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaSignatureValidationToHtml.Run();
        }
    }
}