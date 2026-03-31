using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReference
{
    static void Main()
    {
        // Load the existing XLTX template
        Workbook workbook = new Workbook("Template.xltx");

        // XLTX files do not contain a VBA project.
        // Save as a macro‑enabled workbook to create the VBA project, then reload it.
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (example values)
        // Parameters: name, absoluteLibid, relativeLibid
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            @"C:\AddIns\MyAddIn.xlam",
            @"..\AddIns\MyAddIn.xlam");

        // Save the workbook with the new reference (macro‑enabled)
        workbook.Save("TemplateWithReference.xlsm", SaveFormat.Xlsm);
    }
}