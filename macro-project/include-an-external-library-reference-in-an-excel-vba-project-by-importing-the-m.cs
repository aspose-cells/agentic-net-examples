using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (will become macro‑enabled after saving as .xlsm)
        Workbook workbook = new Workbook();

        // Define the reference details.
        // absoluteLibid – full path to the MHTML reference file.
        // relativeLibid – path relative to the workbook location.
        string referenceName = "MyMhtmlReference";
        string absoluteLibid = @"C:\References\MyAddIn.mht";
        string relativeLibid = @"..\\References\\MyAddIn.mht";

        // Add a reference to the external VBA project (the MHTML file).
        int refIndex = workbook.VbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);
        Console.WriteLine($"Reference added at index {refIndex}. Total references: {workbook.VbaProject.References.Count}");

        // Save the workbook as a macro‑enabled file.
        workbook.Save("WorkbookWithMhtmlReference.xlsm", SaveFormat.Xlsm);
    }
}