using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceDemo
{
    public class AddRegisteredReferenceDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (will become macro-enabled when saved as .xlsm)
                Workbook workbook = new Workbook();

                // Access the VBA project associated with the workbook
                VbaProject vbaProject = workbook.VbaProject;

                // Add a registered reference to an Automation type library (example: stdole)
                vbaProject.References.AddRegisteredReference(
                    "stdole",
                    "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation");

                // Define output file path
                string outputPath = "output.xlsm";

                // Save the workbook as a macro-enabled file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}