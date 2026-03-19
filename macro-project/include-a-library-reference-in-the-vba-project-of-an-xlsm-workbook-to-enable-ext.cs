using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Create a new workbook (initially a regular .xlsx workbook)
        Workbook workbook = new Workbook();

        // Access the VBA project (automatically created for macro-enabled workbooks)
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (e.g., an add‑in)
        // Parameters: reference name, absolute libid (full path), relative libid
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",                              // reference name
            @"C:\AddIns\MyAddIn.xlam",              // absolute path to the add‑in
            @"..\\AddIns\\MyAddIn.xlam");           // relative path (optional)

        // Save the workbook as a macro‑enabled file (.xlsm)
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}