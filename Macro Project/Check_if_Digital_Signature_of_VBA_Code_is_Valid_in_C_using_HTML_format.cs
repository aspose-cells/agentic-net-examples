using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    public class Program
    {
        public static void Main()
        {
            // Load a workbook that contains a VBA project
            Workbook workbook = new Workbook("sample.xlsm"); // load rule

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine signing status
            string signedStatus = vbaProject.IsSigned ? "Signed" : "Not Signed";
            string validStatus = vbaProject.IsSigned
                ? (vbaProject.IsValidSigned ? "Valid" : "Invalid")
                : "N/A";

            // Build a simple HTML report
            string html = "<!DOCTYPE html><html><head><meta charset=\"UTF-8\"><title>VBA Signature Report</title></head><body>";
            html += "<h1>VBA Digital Signature Report</h1>";
            html += $"<p><strong>Workbook:</strong> {Path.GetFileName("sample.xlsm")}</p>";
            html += $"<p><strong>VBA Project Signed:</strong> {signedStatus}</p>";
            html += $"<p><strong>Signature Valid:</strong> {validStatus}</p>";
            html += "</body></html>";

            // Write the HTML report to a file
            File.WriteAllText("VbaSignatureReport.html", html);

            Console.WriteLine("HTML report generated: VbaSignatureReport.html");
        }
    }
}