using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionReport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file path (macro-enabled workbook)
            string inputPath = "input.xlsm";

            // Output HTML report path
            string outputPath = "VbaProtectionReport.html";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine protection status
            bool isProtected = vbaProject.IsProtected;
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Build HTML content
            StringBuilder html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("<meta charset=\"UTF-8\">");
            html.AppendLine("<title>VBA Protection Report</title>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }");
            html.AppendLine("h1 { color: #2E6DA4; }");
            html.AppendLine("table { border-collapse: collapse; width: 50%; }");
            html.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; }");
            html.AppendLine("th { background-color: #f2f2f2; }");
            html.AppendLine("</style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            html.AppendLine("<h1>VBA Protection Report</h1>");
            html.AppendLine("<p><strong>File:</strong> " + Path.GetFileName(inputPath) + "</p>");
            html.AppendLine("<table>");
            html.AppendLine("<tr><th>Property</th><th>Value</th></tr>");
            html.AppendLine("<tr><td>IsProtected</td><td>" + isProtected + "</td></tr>");
            html.AppendLine("<tr><td>IsLockedForViewing</td><td>" + isLockedForViewing + "</td></tr>");
            html.AppendLine("</table>");
            html.AppendLine("</body>");
            html.AppendLine("</html>");

            // Save the HTML report
            File.WriteAllText(outputPath, html.ToString(), Encoding.UTF8);

            // Optional console output
            Console.WriteLine("VBA protection status report generated at: " + outputPath);
        }
    }
}