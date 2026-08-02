// Title: C# – Disable column drag‑and‑drop reordering while allowing width changes with Aspose.Cells worksheet protection
// Description: Creates a new workbook, enables column‑width formatting (AllowFormattingColumn = true), disables column reordering by turning off sorting (AllowSorting = false), applies full protection (ProtectionType.All) without a password, and saves the file as ProtectedWorksheet.xlsx.
// Keywords: Aspose.Cells worksheet protection C# | disable column reordering Aspose.Cells | allow column width changes protection | AllowFormattingColumn example | AllowSorting false Aspose.Cells | ProtectionType.All C#
// Common Searches: Aspose.Cells prevent column drag and drop | C# protect worksheet but allow column resizing | disable sorting in Aspose.Cells worksheet protection | how to keep column order in Aspose.Cells | allow column formatting while protecting sheet Aspose.Cells
// Developer Intent: Protect a worksheet so users can resize columns but cannot change their order.
// Use Cases: Template where column layout must stay fixed while users adjust widths for readability. | Report distribution that preserves column positions for data integrity yet permits column‑width formatting. | Shared workbook that blocks accidental column reordering but allows end‑users to fine‑tune column sizes.
// AI Prompts: Generate C# code with Aspose.Cells that protects a worksheet, allows column width changes, and blocks column drag‑and‑drop reordering. | Show how to set AllowFormattingColumn = true and AllowSorting = false in Aspose.Cells worksheet protection. | Explain configuring Aspose.Cells Protection properties to keep column order fixed while permitting column resizing in a .NET application.

using Aspose.Cells;

// Creates a new workbook, enables column‑width formatting (AllowFormattingColumn = true), disables column reordering by turning off sorting (AllowSorting = false), applies full protection (ProtectionType.All) without a password, and saves the file as ProtectedWorksheet.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Access the worksheet's protection settings
        Protection protection = sheet.Protection;

        // Allow changing column widths (formatting columns)
        protection.AllowFormattingColumn = true;

        // Disable drag‑and‑drop column reordering by disallowing sorting
        protection.AllowSorting = false;

        // Protect the worksheet with all protection options (no password)
        sheet.Protect(ProtectionType.All);

        // Save the protected workbook
        workbook.Save("ProtectedWorksheet.xlsx");
    }
}
