using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddMsXmlReference
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled workbook to ensure a VBA project is created,
        // then reload it so the VbaProject property is available.
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);
        workbook = new Workbook(tempPath);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to Microsoft XML, v6.0 (MSXML6) library
        vbaProject.References.AddRegisteredReference(
            "MSXML2",
            "*\\G{88D969C0-F192-11D4-A65F-0040963251E5}#6.0#0#C:\\Windows\\System32\\msxml6.dll#Microsoft XML, v6.0");

        // Save the workbook with the added reference
        workbook.Save("WorkbookWithMsXmlReference.xlsm", SaveFormat.Xlsm);
    }
}