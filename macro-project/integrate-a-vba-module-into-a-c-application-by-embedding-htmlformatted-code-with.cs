using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaIntegration
{
    public class VbaHtmlModuleDemo
    {
        public static void Run()
        {
            // Create a new macro‑enabled workbook
            Workbook workbook = new Workbook(FileFormatType.Xlsm);

            // Access the VBA project (automatically created for a macro‑enabled workbook)
            VbaProject vbaProject = workbook.VbaProject;

            // Add a new class module to hold the VBA code
            int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "HtmlMacroModule");
            VbaModule vbaModule = vbaProject.Modules[moduleIndex];

            // VBA code that contains an HTML‑formatted string
            string vbaCode =
                "Sub InsertHtml()\n" +
                "    Dim htmlContent As String\n" +
                "    htmlContent = \"<html><body><h1>Welcome to Aspose.Cells</h1>\" & _\n" +
                "                  \"<p>This is generated from VBA.</p></body></html>\"\n" +
                "    ' Example: write the HTML to a temporary file\n" +
                "    Dim fso As Object\n" +
                "    Set fso = CreateObject(\"Scripting.FileSystemObject\")\n" +
                "    Dim tempFile As Object\n" +
                "    Set tempFile = fso.CreateTextFile(\"C:\\\\Temp\\\\generated.html\", True)\n" +
                "    tempFile.Write htmlContent\n" +
                "    tempFile.Close\n" +
                "    MsgBox \"HTML file created at C:\\\\Temp\\\\generated.html\"\n" +
                "End Sub";

            // Assign the VBA code to the module
            vbaModule.Codes = vbaCode;

            // Save the workbook as a macro‑enabled file
            workbook.Save("VbaHtmlModuleDemo.xlsm", SaveFormat.Xlsm);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            VbaHtmlModuleDemo.Run();
        }
    }
}