// Title: Get and Log Worksheet UniqueId (SheetId) with Aspose.Cells for .NET
// Description: Demonstrates how to assign a GUID‑based UniqueId to a worksheet, save the workbook, reload it, retrieve the identifier, and write the worksheet name and UniqueId to the console for audit tracking.
// Keywords: Aspose.Cells | C# | .NET | Worksheet UniqueId | SheetId | audit log | GUID identifier | persist worksheet ID | read worksheet UniqueId | Excel workbook tracking
// Common Searches: Aspose.Cells get worksheet UniqueId | How to log worksheet SheetId in .NET | Persist worksheet identifier after saving | Retrieve worksheet UniqueId for audit | C# Aspose.Cells worksheet ID tracking
// Developer Intent: Extract a worksheet's UniqueId and record it for compliance or change‑tracking purposes.
// Use Cases: Assign a GUID as the UniqueId of a newly created worksheet and store it in the Excel file. | Reload a saved workbook to verify that the persisted UniqueId matches the original value. | Generate an audit report that lists each worksheet name alongside its UniqueId.
// AI Prompts: Write C# code that sets a worksheet's UniqueId to a GUID, saves the workbook, reloads it, and prints the UniqueId for auditing. | Create a reusable method that accepts a workbook path and returns a dictionary of worksheet names and their UniqueIds for compliance reporting. | Explain how Aspose.Cells' UniqueId property can be used to uniquely identify worksheets across sessions and best practices for secure logging.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to assign a GUID‑based UniqueId to a worksheet, save the workbook, reload it, retrieve the identifier, and write the worksheet name and UniqueId to the console for audit tracking.
    public class WorksheetUniqueIdAuditDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (contains a default worksheet)
            Workbook workbook = new Workbook();

            // Generate and assign a unique identifier to the first worksheet
            string generatedId = "{" + Guid.NewGuid().ToString() + "}";
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.UniqueId = generatedId;

            // Define the output file path
            string filePath = "WorksheetUniqueIdAudit.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook so the UniqueId is persisted
            workbook.Save(filePath);

            // Verify the file exists before loading
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"The file '{filePath}' was not found after saving.");
            }

            // Reload the workbook to verify the UniqueId was saved correctly
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

            // Retrieve the UniqueId (acts as the unique SheetId)
            string sheetUniqueId = loadedWorksheet.UniqueId;

            // Log the UniqueId for audit tracking
            Console.WriteLine($"Audit Log - Worksheet Name: {loadedWorksheet.Name}, UniqueId: {sheetUniqueId}");
        }
    }
}
