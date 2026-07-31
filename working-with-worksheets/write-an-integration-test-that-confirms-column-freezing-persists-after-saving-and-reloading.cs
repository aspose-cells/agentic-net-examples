// Title: Integration Test: Verify Freeze Panes Persistence After Save/Load with Aspose.Cells for .NET
// Description: Demonstrates how to create a C# integration test that freezes panes at cell C2, saves the workbook, reloads it, and asserts that the frozen rows and columns are retained using Aspose.Cells' GetFreezedPanes method.
// Keywords: Aspose.Cells freeze panes test | C# Excel freeze panes persistence | GetFreezedPanes integration test | Aspose.Cells save and reload verification | Excel workbook freeze pane unit test | Aspose.Cells .NET testing
// Common Searches: Aspose.Cells verify frozen panes after saving | C# integration test for freeze panes persistence | GetFreezedPanes example after workbook reload | How to test freeze pane settings with Aspose.Cells | Aspose.Cells unit test freeze rows and columns
// Developer Intent: Confirm that frozen rows and columns remain unchanged after a workbook is saved and reopened.
// Use Cases: Automated regression test to ensure freeze pane settings survive file I/O. | CI pipeline validation for reports that require header rows to stay frozen. | Unit testing of grid components that depend on persistent frozen panes.
// AI Prompts: Generate an MSTest method that asserts frozen pane properties after saving and loading an Aspose.Cells workbook. | Convert the freeze pane persistence test to xUnit with proper disposal of temporary files. | Create a PowerShell script to compile the project, execute the freeze pane test, and output a pass/fail result.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to create a C# integration test that freezes panes at cell C2, saves the workbook, reloads it, and asserts that the frozen rows and columns are retained using Aspose.Cells' GetFreezedPanes method.
class FreezePanesIntegrationTest
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze panes at cell C2 (row index 1, column index 2) with 1 frozen row and 2 frozen columns
        sheet.FreezePanes(1, 2, 1, 2);

        // Save the workbook to a temporary file
        string tempFile = Path.Combine(Path.GetTempPath(), "FreezePanesTest.xlsx");
        workbook.Save(tempFile);

        // Load the workbook back from the file
        Workbook loadedWorkbook = new Workbook(tempFile);
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // Retrieve freeze pane information from the loaded worksheet
        bool hasFreeze = loadedSheet.GetFreezedPanes(out int row, out int column, out int frozenRows, out int frozenColumns);

        // Verify that freeze panes were persisted
        if (!hasFreeze)
            throw new Exception("Freeze panes were not persisted after reloading the workbook.");

        if (row != 1 || column != 2 || frozenRows != 1 || frozenColumns != 2)
            throw new Exception($"Freeze pane values mismatch. Expected (1,2,1,2) but got ({row},{column},{frozenRows},{frozenColumns}).");

        Console.WriteLine("Freeze panes persisted correctly after save and reload.");
    }
}
