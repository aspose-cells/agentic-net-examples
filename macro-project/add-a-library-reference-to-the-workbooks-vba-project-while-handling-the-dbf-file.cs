using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceToDbf
{
    static void Main()
    {
        // Load the DBF file into a workbook
        Workbook workbook = new Workbook("sample.dbf");

        // Ensure the workbook has a VBA project (required for adding references)
        if (workbook.VbaProject == null)
        {
            // Save as a macro‑enabled workbook to create the VBA project, then reload
            workbook.Save("temp.xlsm", SaveFormat.Xlsm);
            workbook = new Workbook("temp.xlsm");
            System.IO.File.Delete("temp.xlsm");
        }

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (library)
        // name: reference name
        // absoluteLibid: full path to the referenced project
        // relativeLibid: relative path to the referenced project
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            "C:\\AddIns\\MyAddIn.xlam",
            "..\\AddIns\\MyAddIn.xlam");

        // Save the workbook with the added VBA reference
        workbook.Save("output_with_reference.xlsm", SaveFormat.Xlsm);
    }
}