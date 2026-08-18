// Title: Freeze Top Header Row and Protect Worksheet with Password using Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, freezes the first row by setting the freeze pane at A2, applies full worksheet protection with a password, and saves the file as ProtectedHeader.xlsx.
// Keywords: Aspose.Cells freeze panes C# | Aspose.Cells protect worksheet password | freeze header row C# | worksheet protection Aspose.Cells | C# Excel freeze and lock
// Common Searches: Aspose.Cells freeze first row and protect sheet | C# protect worksheet after freezing panes | how to lock header row in generated Excel using Aspose.Cells | freeze panes and set password protection Aspose.Cells .NET | prevent editing of frozen header in Aspose.Cells workbook
// Developer Intent: Freeze the first row and secure the worksheet with a password to stop accidental header changes.
// Use Cases: Generate a read‑only report where the header stays visible while scrolling and cannot be edited. | Create a data‑entry template that allows users to edit data rows but keeps the frozen header locked. | Distribute a spreadsheet to end users with a protected header after programmatically freezing panes.
// AI Prompts: Show C# code to freeze multiple rows and protect only specific cells while leaving other cells editable with Aspose.Cells. | Explain how to unprotect a worksheet, modify the header row, and re‑apply password protection using Aspose.Cells for .NET. | Compare ProtectionType options when protecting a worksheet after freezing panes in Aspose.Cells.

using System;
using Aspose.Cells;

// Creates a new workbook, freezes the first row by setting the freeze pane at A2, applies full worksheet protection with a password, and saves the file as ProtectedHeader.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze the header row (first row). Freeze starts at cell A2,
        // with 1 frozen row and 0 frozen columns.
        sheet.FreezePanes("A2", 1, 0);

        // Protect the worksheet with all protection options and a password.
        // This prevents editing of the locked header cells.
        sheet.Protect(ProtectionType.All, "HeaderProtect123", null);

        // Save the workbook
        workbook.Save("ProtectedHeader.xlsx");
    }
}
