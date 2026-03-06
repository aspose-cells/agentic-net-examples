using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaProjectReference
{
    static void Main()
    {
        // Create a new workbook (will be macro-enabled after saving as .xlsm)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (example values)
        // name: reference name
        // absoluteLibid: full path to the referenced project
        // relativeLibid: relative path to the referenced project
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",                     // reference name
            @"C:\AddIns\MyAddIn.xlam",     // absolute LIBID
            @"..\\AddIns\\MyAddIn.xlam"    // relative LIBID
        );

        // Save the workbook as a macro-enabled file (XLSM)
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}