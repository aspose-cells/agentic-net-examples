using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSB format supports VBA)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (e.g., an add‑in)
            // Parameters:
            //   name          – logical name for the reference
            //   absoluteLibid – full path identifier of the external project
            //   relativeLibid – relative path identifier (used when the workbook is moved)
            string referenceName = "MyAddIn";
            string absoluteLibid = @"C:\AddIns\MyAddIn.xlam";
            string relativeLibid = @"..\\AddIns\\MyAddIn.xlam";

            vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

            // Save the workbook as an XLSB file (macro‑enabled binary format)
            string outputPath = "WorkbookWithVbaReference.xlsb";
            workbook.Save(outputPath, SaveFormat.Xlsb);

            Console.WriteLine($"Workbook saved to '{outputPath}' with VBA project reference added.");
        }
    }
}