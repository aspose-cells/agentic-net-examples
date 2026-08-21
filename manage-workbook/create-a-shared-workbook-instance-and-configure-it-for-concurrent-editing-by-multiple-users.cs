// Title: Enable Shared Workbook for Concurrent Editing with Aspose.Cells for .NET
// Description: Demonstrates how to create a new Workbook, turn on shared mode (Settings.Shared) for multi‑user editing, optionally protect the shared workbook with a password, save it as an .xlsx file, and reload it to confirm the shared flag remains true.
// Keywords: Aspose.Cells | .NET | shared workbook | concurrent editing | Workbook.Settings.Shared | protect shared workbook password | multi‑user Excel editing | C# Aspose.Cells example | Excel shared mode | Aspose.Cells US developers
// Common Searches: how to enable shared mode in Aspose.Cells .NET | Aspose.Cells protect shared workbook with password | verify shared workbook setting after saving Aspose.Cells | C# create Excel file for concurrent editing Aspose.Cells | Aspose.Cells shared workbook example
// Developer Intent: Create a shared workbook, optionally secure it, and verify the shared configuration for simultaneous editing.
// Use Cases: Generate a new Excel file that multiple users can edit at the same time. | Apply password protection to a shared workbook to limit who can modify it. | Load an existing workbook and programmatically check if it is in shared mode.
// AI Prompts: Show C# code to enable shared mode and add password protection to an Aspose.Cells workbook. | Provide an example that opens a saved shared workbook and reads the Settings.Shared property. | Explain strategies for handling edit conflicts when several users modify a shared workbook using Aspose.Cells.

using Aspose.Cells;
using System;

// Demonstrates how to create a new Workbook, turn on shared mode (Settings.Shared) for multi‑user editing, optionally protect the shared workbook with a password, save it as an .xlsx file, and reload it to confirm the shared flag remains true.
class SharedWorkbookDemo
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Enable shared mode for concurrent editing by multiple users
        workbook.Settings.Shared = true;

        // (Optional) Protect the shared workbook with a password
        // workbook.ProtectSharedWorkbook("myPassword");

        // Save the shared workbook
        string outputPath = "SharedWorkbook.xlsx";
        workbook.Save(outputPath);

        // Load the workbook to verify the shared setting
        Workbook loadedWorkbook = new Workbook(outputPath);
        Console.WriteLine("Shared property value: " + loadedWorkbook.Settings.Shared);
    }
}
