using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project of the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a new class module named "MyModule"
        int moduleIndex = vbaProject.Modules.Add(VbaModuleType.Class, "MyModule");

        // Retrieve the added module and assign VBA code to it
        VbaModule module = vbaProject.Modules[moduleIndex];
        module.Codes = "Sub HelloWorld()\r\n    MsgBox \"Hello from VBA!\"\r\nEnd Sub";

        // Save the workbook in SpreadsheetML (XML) format
        workbook.Save("WorkbookWithVba.xml", SaveFormat.Xml);
    }
}