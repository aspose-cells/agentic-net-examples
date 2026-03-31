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

            // Add a reference to an external VBA project using a PRN‑formatted library file.
            // Parameters:
            //   name          – logical name for the reference
            //   absoluteLibid – full path to the PRN file
            //   relativeLibid – relative path (can be used when the workbook is moved)
            string referenceName = "MyPrnLibrary";
            string absoluteLibid = @"C:\Libraries\MyLibrary.prn";
            string relativeLibid = @"..\Libraries\MyLibrary.prn";

            // Add the project reference
            vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

            // Optional: display the total number of references added
            Console.WriteLine("Total VBA references: " + vbaProject.References.Count);

            // Save the workbook as a macro‑enabled file
            workbook.Save("WorkbookWithPrnReference.xlsm", SaveFormat.Xlsm);
        }
    }
}