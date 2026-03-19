using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class InsertVbaProjectReference
{
    static void Main()
    {
        // Create a new workbook (will be saved as macro‑enabled later)
        Workbook workbook = new Workbook();

        // Access the VBA project associated with the workbook
        VbaProject vbaProject = workbook.VbaProject;

        // Define the reference details (replace with actual paths as needed)
        string referenceName = "MyAddIn";
        string absoluteLibid = @"C:\AddIns\MyAddIn.xlam";          // Absolute path to the add‑in
        string relativeLibid = @"..\\AddIns\\MyAddIn.xlam";       // Relative path to the add‑in

        // Add the external VBA project reference
        int referenceIndex = vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

        // Output confirmation
        Console.WriteLine($"Reference added at index {referenceIndex}. Total references: {vbaProject.References.Count}");

        // Save the workbook as a macro‑enabled file
        workbook.Save("WorkbookWithReference.xlsm", SaveFormat.Xlsm);
    }
}