// Title: Password‑protected read‑only worksheet (table) with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, populates a simple table, and uses Aspose.Cells' Protection object to disable content editing, enable cell selection, set a password, and save the file as a read‑only Excel worksheet.
// Keywords: Aspose.Cells | C# | .NET | worksheet protection | password protection | read only Excel | disable editing cells | allow selecting locked cells | Excel table security | Aspose.Cells protect worksheet
// Common Searches: Aspose.Cells protect worksheet with password C# | make Excel sheet read only using Aspose.Cells | disable editing but allow selection in Aspose.Cells | set worksheet protection options Aspose.Cells .NET | password protect entire table in Excel programmatically
// Developer Intent: Apply a password to make the whole worksheet read‑only while still allowing users to select and view cells.
// Use Cases: Distribute financial reports that must not be altered by recipients. | Publish a template where only designated input cells are unlocked. | Secure a shared spreadsheet before sending it to external stakeholders.
// AI Prompts: Generate C# code with Aspose.Cells to protect a worksheet with a password, disable content editing, and allow selection of both locked and unlocked cells. | Show how to protect only a specific table range while keeping other worksheet areas editable using Aspose.Cells for .NET. | Explain how to verify worksheet protection status after saving an Excel file with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, populates a simple table, and uses Aspose.Cells' Protection object to disable content editing, enable cell selection, set a password, and save the file as a read‑only Excel worksheet.
    public class ProtectWorksheetReadOnly
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Fill sample data to represent the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");

            // Get the protection object of the worksheet
            Protection protection = sheet.Protection;

            // Disallow editing of cell contents (read‑only)
            protection.AllowEditingContent = false;

            // Optionally allow selecting locked/unlocked cells so users can view data
            protection.AllowSelectingLockedCell = true;
            protection.AllowSelectingUnlockedCell = true;

            // Set a password to protect the worksheet
            protection.Password = "ReadOnlyPassword";

            // Save the workbook
            workbook.Save("ProtectedTable.xlsx");
        }
    }
}
