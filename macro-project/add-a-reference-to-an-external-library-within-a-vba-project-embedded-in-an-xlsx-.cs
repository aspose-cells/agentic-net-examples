using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (initially an .xlsx in memory)
            Workbook workbook = new Workbook();

            // Access the VBA project (it is created automatically for macro-enabled workbooks)
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (e.g., an add‑in library)
            // Parameters: reference name, absolute libid (full path), relative libid (relative path)
            vbaProject.References.AddProjectRefrernce(
                "MyExternalLib",
                @"C:\Libraries\MyExternalLib.xlam",
                @"..\Libraries\MyExternalLib.xlam");

            // Save the workbook as a macro‑enabled file so the VBA project (and its reference) is retained
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}