// Title: C# – Protect an Aspose.Cells worksheet while allowing edits only in a named AllowEditRange
// Description: Creates a new workbook, defines an AllowEditRange called "EditableSection" (cells B2:D10), optionally assigns a password to the range, protects the entire worksheet with all protection types, and saves the file as ProtectedEditableSection.xlsx.
// Keywords: Aspose.Cells | C# worksheet protection | AllowEditRange | editable range in protected sheet | Excel API password protection | protect sheet except specific cells | Aspose.Cells Protect method | range‑level security | Excel template locking
// Common Searches: Aspose.Cells protect sheet but keep a range editable | C# AllowEditRange password Aspose.Cells | how to lock all cells except B2:D10 using Aspose.Cells | protect worksheet with editable section Aspose .NET | Aspose.Cells set editable range on protected workbook
// Developer Intent: Secure the entire worksheet while permitting user edits only within the named range "EditableSection".
// Use Cases: Distribute a financial model where calculation cells are locked and input cells are editable. | Create a data‑entry template that restricts users to specific fields while preserving formulas. | Generate a report where only the summary section can be modified after the rest of the sheet is protected.
// AI Prompts: Write C# code with Aspose.Cells to protect a worksheet and define an AllowEditRange named "EditableSection" covering B2:D10, including a password for the range. | Explain how to add multiple AllowEditRanges to a protected worksheet in Aspose.Cells and manage individual passwords. | Show how to programmatically unprotect a sheet, change the coordinates of an existing AllowEditRange, and re‑apply protection using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a new workbook, defines an AllowEditRange called "EditableSection" (cells B2:D10), optionally assigns a password to the range, protects the entire worksheet with all protection types, and saves the file as ProtectedEditableSection.xlsx.
    public class ProtectWorksheetWithEditableRange
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created and saved successfully.");
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

            // Define the area that should remain editable (e.g., B2:D10)
            // Add an allow‑edit range named "EditableSection"
            // Parameters: name, startRow, startColumn, endRow, endColumn (zero‑based indexes)
            int editableRangeIndex = worksheet.AllowEditRanges.Add("EditableSection", 1, 1, 9, 3);
            ProtectedRange editableRange = worksheet.AllowEditRanges[editableRangeIndex];

            // Optional: set a password for the editable range if you want extra security
            // editableRange.Password = "rangePassword";

            // Protect the entire worksheet (all protection types)
            worksheet.Protect(ProtectionType.All);

            // Save the workbook
            workbook.Save("ProtectedEditableSection.xlsx");
        }
    }
}
