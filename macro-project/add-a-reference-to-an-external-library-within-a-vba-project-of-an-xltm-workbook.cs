using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaProjectReference
{
    static void Main()
    {
        // Create a new workbook (macro-enabled template)
        Workbook workbook = new Workbook();

        // Access the VBA project (automatically available for macro-enabled files)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (library)
        // Parameters: name, absoluteLibid (full path), relativeLibid (relative path)
        vbaProject.References.AddProjectRefrernce(
            "MyExternalLib",
            @"C:\Libraries\MyExternalLib.xlam",
            @"..\\MyExternalLib.xlam");

        // Save the workbook as a macro-enabled template (.xltm)
        workbook.Save("WorkbookWithReference.xltm", SaveFormat.Xltm);
    }
}