// Title: Retrieve and Log Worksheet UniqueId with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, accesses the first worksheet, reads its UniqueId property, writes the ID and sheet name to the console for audit tracking, and optionally saves the file. It demonstrates how to use Aspose.Cells for .NET to capture a worksheet's immutable identifier.
// Keywords: Aspose.Cells | C# | .NET | Worksheet UniqueId | get worksheet id | log worksheet identifier | audit tracking | worksheet immutable ID | save workbook Aspose | compliance logging
// Common Searches: How to get a worksheet UniqueId using Aspose.Cells C# | Log worksheet ID for audit with Aspose.Cells .NET | Aspose.Cells UniqueId property example | Retrieve worksheet identifier for compliance reporting | Save workbook after reading worksheet UniqueId
// Developer Intent: Obtain a worksheet's UniqueId and output it for auditing purposes.
// Use Cases: Record the UniqueId when a workbook is generated to maintain an immutable audit trail. | Log the ID before performing bulk edits so changes can be traced back to the original sheet. | Store worksheet UniqueIds in a database for later verification or compliance checks.
// AI Prompts: Generate C# code that iterates through all worksheets in an Aspose.Cells workbook and writes each sheet's UniqueId to a log file. | Explain how to compare stored UniqueId values with current worksheet IDs to detect replacements or renames. | Show how to integrate worksheet UniqueId logging into a .NET microservice for real‑time audit monitoring.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook, accesses the first worksheet, reads its UniqueId property, writes the ID and sheet name to the console for audit tracking, and optionally saves the file. It demonstrates how to use Aspose.Cells for .NET to capture a worksheet's immutable identifier.
    public class WorksheetUniqueIdAudit
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the worksheet's unique identifier
            string uniqueId = worksheet.UniqueId;

            // Log the UniqueId for audit tracking
            Console.WriteLine($"Worksheet '{worksheet.Name}' UniqueId: {uniqueId}");

            // Save the workbook (lifecycle: save) – optional if you need to persist changes
            string outputPath = "WorksheetUniqueIdAudit.xlsx";
            workbook.Save(outputPath);
        }
    }
}
