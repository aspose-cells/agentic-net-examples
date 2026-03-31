using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled when saved as .xlsm)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a project reference.
            // name: logical name of the reference
            // absoluteLibid: full path to the external VBA project (e.g., an .xlam add‑in)
            // relativeLibid: relative path that can be used when the workbook is moved
            vbaProject.References.AddProjectRefrernce(
                "MyExternalAddIn",
                @"C:\AddIns\MyExternalAddIn.xlam",
                @"..\AddIns\MyExternalAddIn.xlam");

            // Optionally display the total number of references added
            Console.WriteLine("Total VBA references: " + vbaProject.References.Count);

            // Save the workbook as a macro‑enabled file so the VBA project is retained
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}