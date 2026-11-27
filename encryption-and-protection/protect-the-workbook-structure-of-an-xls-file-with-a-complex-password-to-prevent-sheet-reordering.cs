// Title: How to protect the structure of an XLS workbook with a complex password using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# protect workbook structure with complex password for XLS files | prevent sheet reordering and deletion in an XLS workbook using Aspose.Cells .NET | set workbook structure protection on legacy XLS format with Aspose.Cells API | C# code to apply structure protection to an existing XLS file using Aspose.Cells | protect Excel 97-2003 workbook structure programmatically with Aspose.Cells
// Tags: Aspose.Cells workbook structure lock | XLS legacy format protection C# | complex password workbook protect Aspose | sheet order protection Aspose.Cells | structure protection API Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing XLS workbook, applies structure protection with a complex password to block sheet reordering, renaming, addition, or deletion, and saves the protected workbook as a new file.
class Program
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Define a complex password for protection
        string password = "C0mpl3x!P@ssw0rd#2026";

        // Protect the workbook structure to prevent sheet reordering, renaming, adding, or deleting
        workbook.Protect(ProtectionType.Structure, password);

        // Save the protected workbook
        workbook.Save("output.xls");
    }
}
