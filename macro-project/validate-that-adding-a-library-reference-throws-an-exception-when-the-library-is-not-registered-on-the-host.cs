// Title: Aspose.Cells .NET: Verify exception when adding an unregistered VBA library reference
// Description: Creates a workbook, accesses its VbaProject, and attempts to add a bogus library via VbaProject.References.AddRegisteredReference. The code catches the expected exception, reports the result, and saves the file as an XLSM workbook.
// Keywords: Aspose.Cells | .NET | C# | VBA | VbaProject | AddRegisteredReference | unregistered library | exception handling | COMException | macro project | Excel XLSM | Windows development | US developers | European .NET community
// Common Searches: Aspose.Cells AddRegisteredReference exception | how to test invalid VBA reference in C# | catch COMException when adding VBA library | validate VBA library registration with Aspose.Cells | unit test for missing type library Aspose.Cells
// Developer Intent: Ensure that calling AddRegisteredReference with a library that is not registered on the host triggers an exception.
// Use Cases: Automated unit test confirming failure for nonexistent VBA type libraries. | Pre‑save validation of macro projects to avoid corrupt workbooks. | User‑friendly error reporting when an invalid reference path is supplied.
// AI Prompts: Write an xUnit test that asserts VbaProject.References.AddRegisteredReference throws a COMException for a fake library ID. | Generate sample code that logs the specific error code returned by AddRegisteredReference when the library is missing. | Explain Aspose.Cells' internal check for registered VBA libraries and list the possible exception types it can raise.

using System;
using Aspose.Cells;
using Aspose.Cells.Vba;

namespace AsposeCellsVbaReferenceValidation
{
    // Creates a workbook, accesses its VbaProject, and attempts to add a bogus library via VbaProject.References.AddRegisteredReference. The code catches the expected exception, reports the result, and saves the file as an XLSM workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided creation rule)
            Workbook workbook = new Workbook();

            // Access the VBA project within the workbook
            VbaProject vbaProject = workbook.VbaProject;

            // Define a bogus library identifier (library not registered on the host)
            string bogusName = "NonExistentLib";
            string bogusLibId = "*\\G{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}#0.0#0#C:\\InvalidPath\\nonexistent.tlb#Invalid";

            bool exceptionThrown = false;

            try
            {
                // Attempt to add the unregistered reference (should throw)
                vbaProject.References.AddRegisteredReference(bogusName, bogusLibId);
            }
            catch (Exception ex)
            {
                // Expected path: an exception is thrown because the library is not registered
                exceptionThrown = true;
                Console.WriteLine("Expected exception caught: " + ex.Message);
            }

            // Validate that the exception was indeed thrown
            if (!exceptionThrown)
            {
                Console.WriteLine("Error: No exception was thrown when adding an unregistered library reference.");
            }
            else
            {
                Console.WriteLine("Validation succeeded: Adding an unregistered library reference throws an exception.");
            }

            // Save the workbook (uses the provided save rule)
            workbook.Save("VbaReferenceValidationResult.xlsm");
        }
    }
}
