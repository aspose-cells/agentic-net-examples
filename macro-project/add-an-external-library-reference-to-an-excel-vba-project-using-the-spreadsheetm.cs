using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (macro-enabled workbook is required for VBA)
            Workbook workbook = new Workbook();

            // Access the VBA project associated with the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Add an external VBA project reference.
            // Parameters:
            //   name          - Logical name of the reference
            //   absoluteLibid - Full path to the external library (absolute)
            //   relativeLibid - Path relative to the workbook location
            vbaProject.References.AddProjectRefrernce(
                "MyExternalLib",
                @"C:\Libraries\MyExternalLib.xlam",
                @"..\Libraries\MyExternalLib.xlam");

            // Save the workbook in SpreadsheetML (XML) format.
            // Note: SaveFormat.SpreadsheetML corresponds to the .xml representation.
            workbook.Save("VbaProjectWithReference.xml", SaveFormat.SpreadsheetML);

            Console.WriteLine("Workbook saved with VBA project reference.");
        }
    }
}