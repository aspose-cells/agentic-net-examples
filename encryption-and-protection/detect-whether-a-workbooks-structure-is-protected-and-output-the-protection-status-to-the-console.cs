// Title: C# – Detect workbook structure protection status using Aspose.Cells for .NET
// Description: Create a workbook, apply structure protection with a password, read the IsProtected flag, output the result, save the file, reload it, and confirm the protection persists.
// Keywords: Aspose.Cells C# workbook protection | check Excel structure protection .NET | Workbook.Settings.IsProtected example | protect workbook structure Aspose | verify protection after save | Aspose.Cells encryption and protection
// Common Searches: Aspose.Cells how to know if workbook structure is protected | C# check Excel workbook protection status | IsProtected property Aspose.Cells | detect workbook structure lock after saving | Aspose.Cells protect workbook structure example
// Developer Intent: Identify whether a workbook's structure is locked and display the boolean result.
// Use Cases: Skip editing steps when the workbook structure is locked. | Audit compliance by confirming protection survives file persistence. | Log protection state of generated Excel files for monitoring.
// AI Prompts: Show C# code that reads Workbook.Settings.IsProtected without saving the workbook. | Demonstrate how to remove structure protection after detecting it with Aspose.Cells. | Compare Workbook.Settings.IsProtected with other protection properties in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsProtectionDemo
{
    // Create a workbook, apply structure protection with a password, read the IsProtected flag, output the result, save the file, reload it, and confirm the protection persists.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "myPassword");

            // Check if the workbook structure (or window) is protected
            bool isStructureProtected = workbook.Settings.IsProtected;

            // Output the protection status to the console
            Console.WriteLine("Workbook structure protected: " + isStructureProtected);

            // Optionally, save the protected workbook (demonstrates lifecycle usage)
            workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

            // Load the saved workbook to verify the protection status persists
            Workbook loadedWorkbook = new Workbook("ProtectedWorkbook.xlsx");
            bool loadedProtectionStatus = loadedWorkbook.Settings.IsProtected;
            Console.WriteLine("Loaded workbook structure protected: " + loadedProtectionStatus);
        }
    }
}
