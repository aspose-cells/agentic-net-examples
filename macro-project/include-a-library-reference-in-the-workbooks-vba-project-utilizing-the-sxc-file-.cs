using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project (created automatically for macro-enabled formats)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (library)
        // Parameters: name, absoluteLibid, relativeLibid
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            @"C:\AddIns\MyAddIn.xlam",
            @"..\AddIns\MyAddIn.xlam");

        // Save the workbook in SXC (OpenOffice Calc) format
        workbook.Save("WorkbookWithReference.sxc", SaveFormat.Sxc);
    }
}