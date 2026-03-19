using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class InsertTxtLibraryReference
{
    static void Main()
    {
        // Create a new workbook (initially without VBA)
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled file to force creation of a VBA project,
        // then reload it so the VbaProject property is available.
        workbook.Save("temp.xlsm", SaveFormat.Xlsm);
        workbook = new Workbook("temp.xlsm");
        System.IO.File.Delete("temp.xlsm");

        // Access the VBA project of the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Define the TXT‑formatted library reference details
        string referenceName = "MyTxtLibrary";
        string absoluteLibid = @"C:\ExternalLibs\MyLibrary.txt";   // full path to the TXT file
        string relativeLibid = @"..\ExternalLibs\MyLibrary.txt";   // relative path from the workbook

        // Add the reference to the VBA project using AddProjectRefrernce
        vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

        // Save the workbook with the new VBA reference
        workbook.Save("WorkbookWithTxtReference.xlsm", SaveFormat.Xlsm);
    }
}