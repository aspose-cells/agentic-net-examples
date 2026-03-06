using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

class AddVbaProjectReference
{
    static void Main()
    {
        // Path to the source workbook (XLSX format)
        string sourcePath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(sourcePath);

        // Ensure the workbook has a VBA project.
        // If not, save as a macro-enabled workbook and reload to create the VBA project.
        if (workbook.VbaProject == null)
        {
            string tempMacroPath = "temp.xlsm";

            // Save as macro-enabled workbook to generate a VBA project
            workbook.Save(tempMacroPath, SaveFormat.Xlsm);

            // Reload the workbook with the VBA project
            workbook = new Workbook(tempMacroPath);

            // Clean up the temporary file
            File.Delete(tempMacroPath);
        }

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Add a reference to an external VBA project (library)
        // Parameters: reference name, absolute libid (full path), relative libid (relative path)
        vbaProject.References.AddProjectRefrernce(
            "MyLibrary",
            @"C:\Addins\MyAddin.xlam",
            @"..\Addins\MyAddin.xlam"
        );

        // Save the workbook back to XLSX format (the VBA project will be stripped,
        // but the reference is added to the VBA project before saving)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine("Reference added and workbook saved to " + outputPath);
    }
}