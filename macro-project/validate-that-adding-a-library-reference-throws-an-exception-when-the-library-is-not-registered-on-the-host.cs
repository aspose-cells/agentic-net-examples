using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the VBA project
            VbaProject vbaProject = workbook.VbaProject;

            // Define a non‑registered library name and libid
            string invalidName = "NonExistentLib";
            string invalidLibId = "*\\G{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}#0.0#0#C:\\InvalidPath\\nonexistent.tlb#Invalid";

            try
            {
                // Attempt to add the unregistered reference (operation under test)
                // This should throw an exception because the library is not registered on the host
                vbaProject.References.AddRegisteredReference(invalidName, invalidLibId);

                // If no exception is thrown, the test has failed
                Console.WriteLine("Test Failed: No exception was thrown when adding an unregistered library reference.");
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown
                Console.WriteLine("Test Passed: Exception caught as expected.");
                Console.WriteLine("Exception Message: " + ex.Message);
            }

            // Save the workbook (lifecycle save) – optional, just to complete the flow
            workbook.Save("VbaReferenceValidationResult.xlsm");
        }
    }
}