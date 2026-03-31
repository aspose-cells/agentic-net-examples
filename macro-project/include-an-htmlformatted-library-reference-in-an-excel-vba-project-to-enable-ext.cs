using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro‑enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add an HTML‑formatted library reference.
            // The libid strings can contain any format required by the external component.
            // Here we use a sample HTML‑styled identifier for demonstration.
            string referenceName = "MyHtmlLib";
            string absoluteLibid = "<html><body>file:///C:/Libraries/MyHtmlLib.xlam</body></html>";
            string relativeLibid = "..\\Libraries\\MyHtmlLib.xlam";

            // Use AddProjectRefrernce to add the external VBA project reference
            vbaProject.References.AddProjectRefrernce(referenceName, absoluteLibid, relativeLibid);

            // Optionally, display the total number of references added
            Console.WriteLine("Total VBA references: " + vbaProject.References.Count);

            // Save the workbook as a macro‑enabled file
            workbook.Save("WorkbookWithHtmlReference.xlsm", SaveFormat.Xlsm);
        }
    }
}