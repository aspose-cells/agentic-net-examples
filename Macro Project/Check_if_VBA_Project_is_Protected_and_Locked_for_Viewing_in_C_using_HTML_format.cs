using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaProtectionCheck
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file (must be a macro-enabled workbook, e.g., .xlsm)
            string inputPath = "sample.xlsm";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Retrieve protection status
            bool isProtected = vbaProject.IsProtected;
            bool isLockedForViewing = vbaProject.IslockedForViewing;

            // Build an HTML report
            string htmlReport = $@"
<html>
<head>
    <title>VBA Project Protection Status</title>
    <style>
        body {{ font-family: Arial, Helvetica, sans-serif; }}
        .status {{ margin-top: 20px; }}
        .status p {{ font-size: 14px; }}
    </style>
</head>
<body>
    <h2>VBA Project Protection Status</h2>
    <div class='status'>
        <p><strong>IsProtected:</strong> {isProtected}</p>
        <p><strong>IsLockedForViewing:</strong> {isLockedForViewing}</p>
    </div>
</body>
</html>";

            // Output HTML to a file
            string outputPath = "VbaProtectionStatus.html";
            File.WriteAllText(outputPath, htmlReport);

            // Optionally, display a confirmation message
            Console.WriteLine($"HTML report generated at: {Path.GetFullPath(outputPath)}");
        }
    }
}