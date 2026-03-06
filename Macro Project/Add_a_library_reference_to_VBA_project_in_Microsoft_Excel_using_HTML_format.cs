using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (e.g., an add‑in)
            // Parameters:
            //   name          – Logical name of the reference
            //   absoluteLibid – Full path identifier of the referenced project
            //   relativeLibid – Relative path identifier (used when the workbook is moved)
            string referenceName = "MyAddIn";
            string absoluteLibid = @"C:\AddIns\MyAddIn.xlam";
            string relativeLibid = @"..\\AddIns\\MyAddIn.xlam";

            vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

            // Optionally display the total number of references added
            Console.WriteLine("Total VBA references: " + vbaProject.References.Count);

            // Save the workbook as a macro‑enabled file (XLSM)
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}