// Title: Aspose.Cells .NET – Disable Row Insertion in a ListObject (Table) via Worksheet Protection
// Description: Creates an Excel workbook, defines a ListObject over a range, and locks the table size by setting worksheet protection AllowInsertingRow to false with a password. The resulting file is a fixed‑size table that end‑users cannot expand.
// Keywords: Aspose.Cells disable row insertion | ListObject protection .NET | prevent adding rows Aspose.Cells | AllowInsertingRow false | Excel table lock Aspose | worksheet protection Aspose.Cells
// Common Searches: how to stop users adding rows to a ListObject in Aspose.Cells | Aspose.Cells set AllowInsertingRow false | protect worksheet to block row insertion .NET | fixed size table Aspose.Cells example | disable table row addition programmatically
// Developer Intent: The developer needs to prevent end‑users from inserting new rows into an Excel ListObject while keeping the rest of the worksheet editable.
// Use Cases: Distribute a template where the data range must remain unchanged. | Enforce data integrity by locking table size in a shared workbook. | Create a reporting sheet that allows edits but forbids row expansion.
// AI Prompts: Show how to lock a ListObject size in Aspose.Cells for .NET without disabling other edits. | Provide code that protects a worksheet and sets AllowInsertingRow to false for a specific table. | Explain how to create a fixed‑size Excel table using Aspose.Cells and prevent row insertion.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates an Excel workbook, defines a ListObject over a range, and locks the table size by setting worksheet protection AllowInsertingRow to false with a password. The resulting file is a fixed‑size table that end‑users cannot expand.
    public class DisableListObjectRowInsertion
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the list object (table)
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Add a ListObject (table) covering the data range A1:B3, with headers
            int listIndex = worksheet.ListObjects.Add("A1", "B3", true);
            ListObject table = worksheet.ListObjects[listIndex];

            // Protect the worksheet and disallow row insertion while protected
            Protection protection = worksheet.Protection;
            protection.AllowInsertingRow = false; // Disable adding new rows
            protection.Password = "securePassword";
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("FixedSizeListObject.xlsx", SaveFormat.Xlsx);
        }
    }
}
