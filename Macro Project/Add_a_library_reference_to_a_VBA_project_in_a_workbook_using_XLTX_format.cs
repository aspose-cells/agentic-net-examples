using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReference
{
    static void Main()
    {
        // Create a new workbook (will be saved as a template later)
        Workbook workbook = new Workbook();

        // Get the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (example paths)
        // Parameters: reference name, absolute libid, relative libid
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            @"C:\Addins\MyAddIn.xlam",
            @"..\Addins\MyAddIn.xlam");

        // Save the workbook as an XLTX template (macro‑free format)
        workbook.Save("WorkbookWithVbaReference.xltx", SaveFormat.Xltx);
    }
}