using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (macro-enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add a reference to an external VBA project (add‑in)
            // Parameters: reference name, absolute libid (full path), relative libid (relative path)
            vbaProject.References.AddProjectRefrernce(
                "MyAddIn",                              // reference name
                @"C:\AddIns\MyAddIn.xlam",              // absolute path to the add‑in
                @"..\AddIns\MyAddIn.xlam");             // relative path to the add‑in

            // Optionally display the total number of references added
            Console.WriteLine("Total VBA references: " + vbaProject.References.Count);

            // Save the workbook as a macro‑enabled file so the VBA project is retained
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}