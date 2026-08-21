// Title: C# Aspose.Cells – Protect worksheet to block column drag‑and‑drop but allow width changes
// Description: Creates a workbook, enables column width editing (AllowFormattingColumn) and disables column insertion, deletion, and sorting. The worksheet is then protected with ProtectionType.All and saved as ProtectedWorksheet.xlsx.
// Keywords: Aspose.Cells worksheet protection C# | disable column drag and drop Aspose.Cells | allow column width changes Aspose.Cells | Prevent column reordering Aspose.Cells | ProtectionType.All column formatting
// Common Searches: Aspose.Cells stop column reordering but keep resize | C# protect worksheet from column insertion deletion sorting | Enable column width editing while disabling drag‑and‑drop in Aspose.Cells | Worksheet.Protect specific AllowFormattingColumn example
// Developer Intent: Apply worksheet protection that blocks column reordering via drag‑and‑drop while still permitting users to adjust column widths.
// Use Cases: Distribute a template where column order must stay fixed but users can resize for readability. | Share a financial report that retains its layout yet lets recipients fit columns to their screens. | Collaborative workbook where structural changes are prohibited but column width customization is allowed.
// AI Prompts: Provide C# Aspose.Cells code to protect a worksheet, disable column insertion, deletion, and sorting, and keep column width editing enabled. | Show an example using ProtectionType.All with AllowFormattingColumn = true and other column actions set to false. | Explain the impact of AllowFormattingColumn, AllowInsertingColumn, AllowDeletingColumn, and AllowSorting when calling sheet.Protect(ProtectionType.All).

using System;
using Aspose.Cells;

// Creates a workbook, enables column width editing (AllowFormattingColumn) and disables column insertion, deletion, and sorting. The worksheet is then protected with ProtectionType.All and saved as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet protection settings
        Protection protection = sheet.Protection;

        // Allow column formatting (e.g., changing column width)
        protection.AllowFormattingColumn = true;

        // Disable column drag‑and‑drop reordering by disallowing insertion,
        // deletion, and sorting of columns
        protection.AllowInsertingColumn = false;
        protection.AllowDeletingColumn = false;
        protection.AllowSorting = false;

        // Apply protection to the worksheet (all protection types)
        sheet.Protect(ProtectionType.All);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
