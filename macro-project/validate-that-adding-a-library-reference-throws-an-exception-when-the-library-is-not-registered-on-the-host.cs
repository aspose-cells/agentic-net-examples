using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

class ValidateUnregisteredLibraryReference
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the VBA project
        VbaProject vbaProject = workbook.VbaProject;

        // Define a library name and libid that are not registered on the host
        string libName = "NonExistentLib";
        string libId = "*\\G{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}#0.0#0#C:\\InvalidPath\\nonexistent.tlb#Invalid";

        bool exceptionThrown = false;

        try
        {
            // Attempt to add the unregistered reference; this should raise an exception
            vbaProject.References.AddRegisteredReference(libName, libId);
        }
        catch (Exception ex)
        {
            // Expected exception caught
            exceptionThrown = true;
            Console.WriteLine("Expected exception caught: " + ex.Message);
        }

        // Verify that the exception was indeed thrown
        if (!exceptionThrown)
        {
            Console.WriteLine("Error: No exception was thrown for an unregistered library reference.");
        }
        else
        {
            Console.WriteLine("Validation succeeded: exception was thrown as expected.");
        }

        // Save the workbook (optional, will contain no references)
        workbook.Save("ValidationResult.xlsm");
    }
}