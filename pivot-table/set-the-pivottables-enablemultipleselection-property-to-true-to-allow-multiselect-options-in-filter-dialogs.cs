// Title: Enable multi‑select in Excel PivotTable filter dialogs using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable multi-select for PivotTable page fields | set EnableMultipleSelection property on PivotField with Aspose.Cells | allow multiple items selection in PivotTable filter using Aspose.Cells .NET | C# code to turn on multi-select in Excel PivotTable filters with Aspose | reflection workaround for EnableMultipleSelection in Aspose.Cells PivotField
// Tags: Aspose.Cells PivotTable multi-select filter | C# EnableMultipleSelection PivotField | Aspose.Cells set page field multiple selection | Excel workbook modify pivot filter options .NET | reflection set PivotField property Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, locates the first PivotTable, iterates its page fields, and uses reflection to set each field's EnableMultipleSelection property to true, enabling multi‑select in filter dialogs before saving the workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one PivotTable
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first PivotTable
                PivotTable pivot = sheet.PivotTables[0];

                // Enable multiple selection for all page fields (filter fields) using reflection
                foreach (PivotField pageField in pivot.PageFields)
                {
                    try
                    {
                        var prop = pageField.GetType().GetProperty("EnableMultipleSelection");
                        if (prop != null && prop.CanWrite)
                        {
                            prop.SetValue(pageField, true);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to enable multiple selection for field '{pageField.Name}': {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
