using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceToOds
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project (created automatically for macro-enabled workbooks)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project
        // Parameters: name, absoluteLibid, relativeLibid
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            @"C:\AddIns\MyAddIn.xlam",
            @"..\AddIns\MyAddIn.xlam");

        // Save the workbook in ODS format
        workbook.Save("WorkbookWithVbaReference.ods", SaveFormat.Ods);
    }
}