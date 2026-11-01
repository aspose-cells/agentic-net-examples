// Title: Collapse all PivotItems in an Excel PivotTable using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# collapse all pivot items in a pivot table | set PivotItem IsVisible false for all rows and columns using Aspose.Cells | programmatically hide pivot field items in Excel with Aspose.Cells .NET | iterate PivotFields to collapse items in an existing workbook using C# | how to collapse pivot table rows and columns via Aspose.Cells API
// Tags: collapse pivot table items Aspose.Cells | set PivotItem IsVisible false C# | hide pivot field items programmatically | modify Excel pivot table visibility .NET | iterate PivotFields Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCollapse
{
    // The example loads an existing workbook, iterates through each PivotTable on the first worksheet, loops over all row and column PivotFields, and uses a dynamic cast to set each PivotItem's IsVisible property to false, collapsing the items, then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Access the collection of pivot tables on the worksheet
                PivotTableCollection pivotTables = sheet.PivotTables;

                // Iterate through each pivot table
                foreach (PivotTable pivotTable in pivotTables)
                {
                    // Collapse items in row fields
                    foreach (PivotField rowField in pivotTable.RowFields)
                    {
                        foreach (PivotItem item in rowField.PivotItems)
                        {
                            try
                            {
                                // Use dynamic to set IsVisible at runtime (avoids compile‑time binding issues)
                                dynamic dynItem = item;
                                dynItem.IsVisible = false;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Unable to hide row item '{item.Name}'. {ex.Message}");
                            }
                        }
                    }

                    // Collapse items in column fields (optional)
                    foreach (PivotField columnField in pivotTable.ColumnFields)
                    {
                        foreach (PivotItem item in columnField.PivotItems)
                        {
                            try
                            {
                                dynamic dynItem = item;
                                dynItem.IsVisible = false;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Warning: Unable to hide column item '{item.Name}'. {ex.Message}");
                            }
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
