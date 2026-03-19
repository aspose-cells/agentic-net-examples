using System;
using Aspose.Cells;
using Aspose.Cells.Vba;
using System.Text;

class ProtectVbaProject
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled file and reload to ensure a VBA project exists
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);
        workbook = new Workbook(tempPath);
        System.IO.File.Delete(tempPath);

        // Add a simple VBA module (optional, just to have content)
        int moduleIndex = workbook.VbaProject.Modules.Add(VbaModuleType.Class, "DemoModule");
        VbaModule module = workbook.VbaProject.Modules[moduleIndex];
        module.Codes = "Sub Demo()\r\n    MsgBox \"Hello from VBA\"\r\nEnd Sub";

        // Protect the VBA project with a password and lock it for viewing
        workbook.VbaProject.Protect(true, "MySecretPassword");

        // Save the workbook as an XLTX template
        workbook.Save("ProtectedTemplate.xltx", SaveFormat.Xltx);
    }
}