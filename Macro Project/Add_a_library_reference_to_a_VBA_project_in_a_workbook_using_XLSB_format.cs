using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (initially without any worksheets)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (e.g., an add‑in)
            // Parameters: reference name, absolute libid (full path), relative libid (relative path)
            vbaProject.References.AddProjectRefrernce(
                "MyAddIn",                              // Reference name
                @"C:\Addins\MyAddIn.xlam",              // Absolute path to the external VBA project
                @"..\Addins\MyAddIn.xlam");             // Relative path (optional)

            // Save the workbook in XLSB format (macro‑enabled binary workbook)
            workbook.Save("WorkbookWithReference.xlsb", SaveFormat.Xlsb);
        }
    }
}