using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class Program
{
    static void Main()
    {
        // Load the OXPS file (replace with actual path)
        Workbook workbook = new Workbook("input.oxps");

        // Ensure the workbook has a VBA project.
        // If the workbook is not macro-enabled, save it as .xlsm and reload.
        if (workbook.VbaProject == null)
        {
            string tempPath = "temp.xlsm";
            workbook.Save(tempPath, SaveFormat.Xlsm);
            workbook = new Workbook(tempPath);
            System.IO.File.Delete(tempPath);
        }

        // Access the VBA project.
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project.
        // Parameters: reference name, absolute libid, relative libid.
        vbaProject.References.AddProjectRefrernce(
            "MyAddIn",
            @"C:\AddIns\MyAddIn.xlam",
            @"..\AddIns\MyAddIn.xlam");

        // Save the workbook with the new reference.
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}