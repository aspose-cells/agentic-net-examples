// Title: Create a new Aspose.Cells workbook and add a single named worksheet in C#
// AI Prompts: Write C# code that creates an Aspose.Cells Workbook, removes any pre‑existing worksheets, and inserts a worksheet named "Sheet1". | Show a reusable C# method that returns a Workbook containing exactly one worksheet with a custom name using Aspose.Cells. | Demonstrate how to clear the default sheets of a newly created workbook before adding a single sheet in Aspose.Cells for C#.
// Common Searches: Aspose.Cells C# create workbook without default sheets | how to add only one worksheet to a new workbook using Aspose.Cells | remove default worksheets from Aspose.Cells workbook before adding new sheet | C# Aspose.Cells initialize workbook with a single named sheet | Aspose.Cells clear all worksheets then add a specific sheet in C#
// Tags: add single worksheet Aspose.Cells C# | clear default worksheets Aspose.Cells | initialize empty workbook Aspose.Cells | custom sheet name Aspose.Cells workbook | delete all worksheets Aspose.Cells

using Aspose.Cells;

// // Creates a new Aspose.Cells Workbook, clears any default worksheets, and adds a single worksheet named "Sheet1".
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Remove any default worksheets (optional, ensures only one sheet)
        workbook.Worksheets.Clear();

        // Add a single worksheet named "Sheet1"
        Worksheet sheet = workbook.Worksheets.Add("Sheet1");

        // The workbook now contains exactly one worksheet.
        // You can save the workbook if needed, e.g.:
        // workbook.Save("MyWorkbook.xlsx");
    }
}
