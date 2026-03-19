using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionInfo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input Excel file (must be macro-enabled, e.g., .xlsm)
            string inputPath = "sample.xlsm";
            // Output HTML file that will contain the protection information
            string outputHtmlPath = "VbaProtectionInfo.html";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Determine protection status
            bool isProtected = vbaProject.IsProtected;
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Build a simple HTML report
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""UTF-8"">
    <title>VBA Project Protection Info</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h2 {{ color: #2F4F4F; }}
        .status {{ margin-top: 10px; }}
        .status p {{ font-size: 1.1em; }}
    </style>
</head>
<body>
    <h2>VBA Project Protection Status</h2>
    <div class=""status"">
        <p><strong>Is Protected:</strong> {isProtected}</p>
        <p><strong>Is Locked For Viewing:</strong> {isLockedForViewing}</p>
    </div>
</body>
</html>";

            // Write the HTML to the output file
            File.WriteAllText(outputHtmlPath, htmlContent);

            Console.WriteLine($"Protection information saved to: {outputHtmlPath}");
        }
    }
}