// Title: Hide confidential row and column items in an Excel pivot table with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide specific row items in a pivot table | How to exclude confidential regions from an Excel pivot using Aspose.Cells | Programmatically hide column field items in a pivot table with Aspose.Cells .NET | Set PivotItem.IsHidden to true in Aspose.Cells example | Hide pivot items based on name in C# Aspose.Cells workbook
// Tags: Aspose.Cells hide pivot row items C# | Aspose.Cells hide pivot column items .NET | set PivotItem.IsHidden property Aspose.Cells | exclude confidential data from Excel pivot Aspose.Cells | programmatic pivot item visibility C# | Aspose.Cells pivot table item filtering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotItemHider
{
    // The example loads an existing workbook, accesses the first pivot table, and hides selected row and column items (e.g., "ConfidentialRegion" and "ClassifiedProduct") by setting each PivotItem's IsHidden property to true, then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook that contains the pivot table.
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust if needed).
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table.
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Retrieve the first pivot table.
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Hide specific row field items (e.g., confidential regions).
                PivotField rowField = pivotTable.RowFields["Region"];
                if (rowField != null)
                {
                    // Use PivotItems collection to access PivotItem objects.
                    foreach (PivotItem item in rowField.PivotItems)
                    {
                        if (item.Name == "ConfidentialRegion")
                        {
                            item.IsHidden = true;
                        }
                        // Additional conditions can be added here.
                    }
                }

                // Hide specific column field items (e.g., classified products).
                PivotField columnField = pivotTable.ColumnFields["Product"];
                if (columnField != null)
                {
                    foreach (PivotItem item in columnField.PivotItems)
                    {
                        if (item.Name == "ClassifiedProduct")
                        {
                            item.IsHidden = true;
                        }
                    }
                }

                // Save the modified workbook.
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to {outputPath}");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
