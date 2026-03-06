using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled workbook will be created on save)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (library)
            // Parameters: reference name, absolute libid (full path), relative libid (relative path)
            vbaProject.References.AddProjectRefrernce(
                "MyLibrary",
                @"C:\Libraries\MyLibrary.xlam",
                @"..\Libraries\MyLibrary.xlam");

            // Save the workbook in XLSM format to preserve the VBA project and its references
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}