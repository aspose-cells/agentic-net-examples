// Title: Load an Excel workbook and retrieve the first worksheet by zero‑based index in C# with Aspose.Cells
// AI Prompts: Write C# code that opens an existing .xlsx file using Aspose.Cells, accesses the worksheet at index 0, assigns it to a Worksheet variable, and then saves any changes. | Show how to rename a worksheet after obtaining it by zero‑based index and persist the workbook using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# get worksheet at index 0 from workbook | how to access first sheet in an Excel file using Aspose.Cells .NET | retrieve worksheet by zero based index Aspose.Cells example | rename worksheet after loading workbook with Aspose.Cells C# | save changes after modifying worksheet using Aspose.Cells
// Tags: load workbook and access worksheet by index Aspose.Cells | zero‑based worksheet collection indexing C# | rename worksheet after retrieval Aspose.Cells | save modified workbook Aspose.Cells .NET | worksheet variable assignment Aspose.Cells C#

using Aspose.Cells;

// Loads an existing workbook, accesses the first worksheet via its zero‑based index, optionally renames it, and saves the workbook using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet at zero‑based index 0 and assign it to a variable
        Worksheet worksheet = workbook.Worksheets[0];

        // (Optional) Perform operations on the worksheet, e.g., rename it
        worksheet.Name = "FirstSheet";

        // Save the workbook to persist any changes (replace the path as needed)
        workbook.Save("output.xlsx");
    }
}
