using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (no worksheets are required for this demo)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (add‑in)
            // Parameters: reference name, absolute libid (full path), relative libid (relative path)
            vbaProject.References.AddProjectRefrernce(
                "MyAddInReference",
                @"C:\Addins\MyAddIn.xlam",
                @"..\Addins\MyAddIn.xlam");

            // Save the workbook in FODS (Flat OpenDocument Spreadsheet) format
            workbook.Save("WorkbookWithVbaReference.fods", SaveFormat.Fods);
        }
    }
}