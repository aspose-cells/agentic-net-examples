using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaReferenceAndSavePdf
{
    static void Main()
    {
        // Create a new workbook (default is .xlsx)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (add-in)
        // Parameters: reference name, absolute libid (full path), relative libid (relative path)
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",                         // reference name
            @"C:\AddIns\MyAddIn.xlam",         // absolute libid
            @"..\AddIns\MyAddIn.xlam");        // relative libid

        // Save the workbook as PDF
        workbook.Save("WorkbookWithVbaReference.pdf", SaveFormat.Pdf);
    }
}