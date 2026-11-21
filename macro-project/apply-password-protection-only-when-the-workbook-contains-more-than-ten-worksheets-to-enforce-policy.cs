// Title: Apply password protection to an Excel workbook's structure with Aspose.Cells for .NET only when it contains more than ten worksheets
// Common Searches: C# Aspose.Cells protect workbook structure with password if worksheet count > 10 | How to conditionally set Excel file password based on number of sheets using Aspose.Cells | Aspose.Cells .NET example for applying workbook protection only when more than ten worksheets exist | Conditional workbook security policy in C# with Aspose.Cells and worksheet threshold | Protect Excel workbook structure programmatically when sheet count exceeds ten in .NET
// Tags: Aspose.Cells workbook structure protection | C# conditional Excel password protection | protect workbook when sheet count exceeds limit | Excel security policy enforcement with Aspose.Cells | apply password to workbook structure programmatically

using Aspose.Cells;

// The code loads an Excel workbook, checks if it has more than ten worksheets, and if so protects the workbook's structure with a password using Aspose.Cells before saving the file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply password protection only if there are more than ten worksheets
        if (workbook.Worksheets.Count > 10)
        {
            // Protect the workbook structure with a password
            workbook.Protect(ProtectionType.Structure, "StrongPassword123");
        }

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
