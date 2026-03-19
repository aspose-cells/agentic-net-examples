using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddPdfLibraryReference
{
    static void Main()
    {
        // Create a new workbook (initially without VBA project)
        Workbook workbook = new Workbook();

        // Save as a macro‑enabled workbook to create the VBA project, then reload it
        string tempFile = "temp.xlsm";
        workbook.Save(tempFile, SaveFormat.Xlsm);
        workbook = new Workbook(tempFile);
        File.Delete(tempFile);

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Define PDF library reference details (adjust paths as needed)
        string referenceName = "PDFLib";
        string absoluteLibid = @"C:\Libraries\PDFLib.tlb";          // absolute path to the type library
        string relativeLibid = @"..\Libraries\PDFLib.tlb";          // relative path from the workbook

        // Add the PDF library reference to the VBA project
        vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

        // Save the workbook with the new reference
        workbook.Save("WorkbookWithPdfReference.xlsm", SaveFormat.Xlsm);
    }
}