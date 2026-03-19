using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaSignatureCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (macro-enabled workbook)
            string workbookPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine if the VBA project is digitally signed
            bool isSigned = vbaProject != null && vbaProject.IsSigned;

            // Build a simple HTML report
            string htmlContent = "<!DOCTYPE html>\n" +
                                 "<html>\n" +
                                 "<head><title>VBA Signature Check</title></head>\n" +
                                 "<body>\n" +
                                 $"<h1>VBA Project Signature Status</h1>\n" +
                                 $"<p>Workbook: {Path.GetFileName(workbookPath)}</p>\n" +
                                 $"<p>VBA Project Signed: {(isSigned ? "Yes" : "No")}</p>\n";

            // If signed, also display whether the signature is valid
            if (isSigned)
            {
                htmlContent += $"<p>Signature Valid: {(vbaProject.IsValidSigned ? "Yes" : "No")}</p>\n";
            }

            htmlContent += "</body>\n</html>";

            // Output the HTML to a file
            string outputPath = "VbaSignatureReport.html";
            File.WriteAllText(outputPath, htmlContent);

            Console.WriteLine($"HTML report generated at: {Path.GetFullPath(outputPath)}");
        }
    }
}