using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceFromMht
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled workbook to ensure a VBA project exists,
        // then reload it so that workbook.VbaProject is not null.
        string tempFile = Path.Combine(Path.GetTempPath(), "temp.xlsm");
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);
        File.Delete(tempFile);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Path to the MHT‑formatted reference file
        string mhtFilePath = @"C:\References\MyLibrary.mht";

        // Define reference name and libids (absolute and relative)
        string referenceName = "MyLibrary";
        string absoluteLibid = mhtFilePath;                                 // full path
        string relativeLibid = @"..\\References\\MyLibrary.mht";            // example relative path

        // Add the external VBA project reference using the MHT file
        vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

        // Save the workbook with the added reference
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}