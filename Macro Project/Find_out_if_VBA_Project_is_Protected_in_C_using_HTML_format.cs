using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class CheckVbaProjectProtection
{
    static void Main(string[] args)
    {
        // Path to the macro-enabled workbook; can be passed as a command‑line argument
        string workbookPath = args.Length > 0 ? args[0] : "sample.xlsm";

        // Load the workbook (Aspose.Cells handles .xlsm files with VBA projects)
        Workbook workbook = new Workbook(workbookPath);

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Retrieve the protection status of the VBA project
        bool isProtected = vbaProject.IsProtected;

        // Create a simple HTML report showing the result
        string htmlContent = $@"
<html>
<head><title>VBA Project Protection Status</title></head>
<body>
    <h2>VBA Project Protection Status</h2>
    <p><strong>File:</strong> {Path.GetFileName(workbookPath)}</p>
    <p><strong>Is Protected:</strong> {isProtected}</p>
</body>
</html>";

        // Save the HTML report next to the original workbook
        string htmlPath = Path.ChangeExtension(workbookPath, ".html");
        File.WriteAllText(htmlPath, htmlContent);

        Console.WriteLine($"Protection status written to: {htmlPath}");
    }
}