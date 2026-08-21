// Title: C# – Add Password Protection to a Shared Excel Workbook with Aspose.Cells
// Description: Demonstrates how to load an existing Excel file, enable shared mode, secure the shared workbook with a password using the ProtectSharedWorkbook method, and write the protected version to a new file.
// Keywords: Aspose.Cells C# protect shared workbook | password protect shared Excel file | ProtectSharedWorkbook method | shared mode encryption Aspose | save protected workbook .NET
// Common Searches: how to password‑lock a shared workbook using Aspose.Cells | Aspose.Cells C# protect shared mode programmatically | set shared workbook password in .NET | save Excel file with shared protection Aspose
// Developer Intent: The developer needs to open a pre‑existing workbook, ensure it is in shared mode, apply a password to that shared state, and persist the secured file.
// Use Cases: Allow multiple users to edit a workbook while requiring a password to modify the shared settings. | Distribute a collaborative Excel file internally but block unauthorized changes to the shared configuration. | Automate creation of shared workbooks that must be unlocked with a secret before users can join the editing session.
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, activates shared mode, applies a password via ProtectSharedWorkbook, and saves the result. | Explain the parameters and behavior of ProtectSharedWorkbook, including how it differs from regular workbook protection. | Provide robust error handling for loading a file, checking shared status, applying password protection, and saving the protected workbook.

using System;
using Aspose.Cells;

// Demonstrates how to load an existing Excel file, enable shared mode, secure the shared workbook with a password using the ProtectSharedWorkbook method, and write the protected version to a new file.
class Program
{
    static void Main()
    {
        // Load the existing shared workbook
        Workbook workbook = new Workbook("SharedWorkbook.xlsx");

        // Ensure the workbook is marked as shared (optional if already shared)
        workbook.Settings.Shared = true;

        // Apply password protection to the shared mode
        workbook.ProtectSharedWorkbook("MySecretPassword");

        // Save the protected workbook
        workbook.Save("ProtectedSharedWorkbook.xlsx");
    }
}
