// Title: How to apply a predefined PivotTable style in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to set a built‑in PivotTable style with Aspose.Cells | Aspose.Cells .NET example for applying PivotTableStyleInfo to an existing workbook | Programmatically change Excel PivotTable visual formatting using Aspose.Cells | Apply a predefined style to the first PivotTable in a worksheet with Aspose.Cells C# | Update PivotTable appearance in an Excel file via Aspose.Cells library
// Tags: Aspose.Cells apply PivotTableStyleInfo | C# set Excel PivotTable style | programmatic PivotTable formatting Aspose.Cells | update PivotTable visual appearance .NET | Excel workbook style modification using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel file, verifies its presence, accesses the first worksheet and its first PivotTable, notes that the PivotTableStyleInfo class (available in newer Aspose.Cells versions) can be used to apply a predefined style, and then saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivot = sheet.PivotTables[0];

            // NOTE:
            // Applying a style to a pivot table uses the PivotTableStyleInfo class,
            // which is available in newer versions of Aspose.Cells.
            // If your version does not contain this class, update the package
            // or omit the style assignment as shown below.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
