// Title: Protect Excel Workbook Structure with a Complex Password using Aspose.Cells for .NET
// Description: Shows how to create or load a workbook, apply structure‑only protection with a strong password via Workbook.Protect, block sheet reordering, addition or deletion, and save the result as an .xlsx file.
// Keywords: Aspose.Cells | Workbook.Protect | structure protection | complex password | C# Excel security | prevent sheet reordering | Excel workbook protection .NET
// Common Searches: Aspose.Cells protect workbook structure C# | set complex password for Excel workbook .NET | prevent sheet reordering Aspose.Cells | structure only protection Excel file | lock workbook structure programmatically
// Developer Intent: Apply structure‑only protection to an Excel workbook with a robust password using Aspose.Cells for .NET.
// Use Cases: Distribute a template workbook that cannot be altered, added to, or have its sheets rearranged. | Secure automatically generated reports so recipients can view data but cannot modify sheet order or create new worksheets. | Enforce integrity in batch export pipelines by locking the workbook structure with a strong password.
// AI Prompts: Write C# code that uses Aspose.Cells to protect only the workbook structure with a given complex password and saves the file as .xlsx. | Show how to verify whether a workbook's structure is protected and retrieve its protection settings using Aspose.Cells for .NET. | Provide an example of unprotecting a workbook structure by supplying the same complex password with Aspose.Cells.

using System;
using Aspose.Cells;

// Shows how to create or load a workbook, apply structure‑only protection with a strong password via Workbook.Protect, block sheet reordering, addition or deletion, and save the result as an .xlsx file.
class ProtectWorkbookStructure
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Define a complex password to protect the workbook structure
        string complexPassword = "P@55w0rd!#2026$%^&*()_+|~`";

        // Protect only the workbook structure (prevents sheet reordering, addition, deletion)
        workbook.Protect(ProtectionType.Structure, complexPassword);

        // Save the protected workbook
        workbook.Save("ProtectedWorkbook.xlsx", SaveFormat.Xlsx);

        // Release resources
        workbook.Dispose();
    }
}
