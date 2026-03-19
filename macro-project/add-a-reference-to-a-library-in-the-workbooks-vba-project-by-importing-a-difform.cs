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

            // Add a reference to an external VBA project (e.g., an add‑in library).
            // The absoluteLibid can be a full path to the .xlam file,
            // and the relativeLibid can be a relative path used by Excel.
            // Adjust the paths to point to your actual library file.
            string absoluteLibid = @"C:\Libraries\MyAddIn.xlam";
            string relativeLibid = @"..\\MyAddIn.xlam";

            // Add the project reference; the method returns the index of the new reference.
            int referenceIndex = vbaProject.References.AddProjectRefrernce("MyAddIn", absoluteLibid, relativeLibid);

            // Optional: display information about the added reference
            VbaProjectReference addedReference = vbaProject.References[referenceIndex];
            Console.WriteLine($"Reference added: Name={addedReference.Name}, AbsoluteLibid={addedReference.Libid}, RelativeLibid={addedReference.RelativeLibid}");

            // Save the workbook as a macro‑enabled file so the VBA project (and its references) are retained
            workbook.Save("WorkbookWithVbaReference.xlsm", SaveFormat.Xlsm);
        }
    }
}