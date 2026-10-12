// Title: Duplicate an Excel worksheet while preserving its pivot tables with Aspose.Cells for .NET
// AI Prompts: Create a copy of a specific worksheet in a workbook, ensuring all pivot tables are retained, using Aspose.Cells in C#. | Generate a renamed duplicate of a sheet that includes its pivot tables and save the result with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells copy worksheet with pivot tables C# example | How to retain pivot tables when duplicating a sheet using Aspose.Cells | C# duplicate Excel sheet preserving pivot cache Aspose.Cells | AddCopy method to clone worksheet and keep pivot tables in .NET
// Tags: Aspose.Cells AddCopy method usage | copy worksheet with pivot tables .NET | retain pivot cache during sheet duplication | rename duplicated worksheet Aspose.Cells | clone Excel sheet preserving pivot tables

using Aspose.Cells;
using System;

// // Loads a workbook, duplicates a specified worksheet (including its pivot tables) using AddCopy, renames the new sheet, and saves the workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Identify the worksheet to duplicate (by name or index)
        Worksheet sourceSheet = workbook.Worksheets["Sheet1"]; // replace with your sheet name

        // Duplicate the worksheet; this also copies any pivot tables it contains
        int copiedIndex = workbook.Worksheets.AddCopy(sourceSheet.Index);

        // Optionally rename the copied worksheet for clarity
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
        copiedSheet.Name = sourceSheet.Name + "_Copy";

        // Save the workbook with the duplicated worksheet and its pivot tables retained
        workbook.Save("output.xlsx");
    }
}
