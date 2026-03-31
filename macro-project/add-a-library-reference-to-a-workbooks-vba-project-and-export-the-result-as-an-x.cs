using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceAndExportXps
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Ensure the workbook has a VBA project by saving as a macro-enabled file and reloading it
        string tempPath = "temp.xlsm";
        workbook.Save(tempPath, SaveFormat.Xlsm);
        workbook = new Workbook(tempPath);
        File.Delete(tempPath);

        // Add a reference to an external VBA project
        // Parameters: name, absoluteLibid, relativeLibid
        workbook.VbaProject.References.AddProjectRefrernce(
            "MyReference",
            "absoluteLibid",
            "relativeLibid");

        // Export the workbook (with VBA project) as an XPS document
        workbook.Save("WorkbookWithVbaReference.xps", SaveFormat.Xps);
    }
}