// Title: Protect an XLS workbook’s structure with a strong password using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xls file, applies workbook structure protection with a complex password via Aspose.Cells, and saves the protected file. | Show how to call Aspose.Cells Protect method to block sheet reordering, addition, deletion, and renaming in an XLS workbook using a custom password.
// Common Searches: Aspose.Cells C# protect workbook structure with password for .xls files | How to prevent sheet reordering in an old Excel (XLS) using Aspose.Cells | Set complex password for workbook structure protection in Aspose.Cells .NET | C# example to lock Excel workbook structure and disable sheet changes with Aspose.Cells | Apply structure protection to XLS workbook using Aspose.Cells Protect method
// Tags: Aspose.Cells workbook structure protection C# | XLS file password protection Aspose.Cells | block sheet order changes Aspose.Cells | complex password workbook protect method | Protect(ProtectionType.Structure) usage .NET

using Aspose.Cells;

// The sample loads an existing XLS workbook, calls workbook.Protect with ProtectionType.Structure and a strong password to disable sheet reordering, addition, deletion, and renaming, then saves the protected file as output.xls.
class WorkbookProtectionExample
{
    static void Main()
    {
        // Load the existing XLS workbook
        Workbook workbook = new Workbook("input.xls");

        // Protect the workbook structure with a complex password
        // This prevents sheet reordering, adding, deleting, or renaming
        string complexPassword = "C0mpl3xP@ssw0rd!#2026";
        workbook.Protect(ProtectionType.Structure, complexPassword);

        // Save the protected workbook
        workbook.Save("output.xls");
    }
}
