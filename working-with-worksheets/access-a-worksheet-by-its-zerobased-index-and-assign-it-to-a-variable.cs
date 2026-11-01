// Title: How to get a worksheet by zero‑based index and assign it to a variable with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that accesses the first worksheet of a Workbook using its zero‑based index and stores it in a Worksheet variable. | Show an example of renaming a worksheet after retrieving it by numeric index with Aspose.Cells. | Explain how to loop through a workbook’s Worksheets collection by index positions in C#.
// Common Searches: Aspose.Cells C# retrieve worksheet at index 0 | how to assign worksheet to variable using Aspose.Cells .NET | rename first sheet after accessing by numeric index in Aspose.Cells | C# get worksheet from workbook by zero based index Aspose.Cells | access worksheet collection by index position Aspose.Cells .NET
// Tags: worksheet index access Aspose.Cells | assign worksheet variable C# | rename worksheet after index retrieval | Aspose.Cells workbook worksheets collection | zero based worksheet indexing .NET

using Aspose.Cells;

// Creates a new Workbook, accesses the first worksheet via its zero‑based index, assigns it to a Worksheet variable, and renames it to "FirstSheet".
class Program
{
    static void Main()
    {
        // Create a new workbook (you could also load an existing file)
        Workbook workbook = new Workbook();

        // Access the worksheet at zero‑based index 0 and assign it to a variable
        Worksheet worksheet = workbook.Worksheets[0];

        // Example usage: rename the worksheet
        worksheet.Name = "FirstSheet";
    }
}
