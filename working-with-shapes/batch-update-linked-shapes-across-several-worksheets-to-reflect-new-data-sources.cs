// Title: C# – Batch update linked shapes and external data sources with Aspose.Cells
// Description: Loads a primary workbook, replaces its linked data sources with specified external workbooks, refreshes every shape (dropdowns, list boxes, etc.) on all worksheets using UpdateSelectedValue, recalculates formulas, and saves the updated file.
// Keywords: Aspose.Cells | C# | linked shapes | external data source | refresh dropdown | batch shape update | Excel workbook automation | UpdateSelectedValue | recalculate formulas | .NET Excel example | GitHub sample
// Common Searches: Aspose.Cells update linked shapes C# | Refresh Excel dropdowns programmatically | Batch replace external data source in workbook | How to recalculate formulas after linked data change | C# code to update linked shapes across worksheets
// Developer Intent: Refresh all linked shapes to reflect new external workbooks and ensure dependent formulas are recalculated.
// Use Cases: Automatically refresh dropdown lists that reference external workbooks after source files are modified. | Synchronize list box selections across multiple sheets when data sources change. | Maintain data integrity by recalculating formulas after batch updating linked shapes. | Prepare a workbook for distribution with updated linked data without opening Excel.
// AI Prompts: Generate C# Aspose.Cells code that loads a main workbook, replaces its linked data sources with an array of external workbooks, updates every shape on all worksheets, recalculates formulas, and saves the file. | Explain error handling for shapes that do not support UpdateSelectedValue in a batch refresh scenario. | Show how to verify that each linked shape’s source range matches the new external workbook before saving.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace LinkedShapeBatchUpdate
{
    // Loads a primary workbook, replaces its linked data sources with specified external workbooks, refreshes every shape (dropdowns, list boxes, etc.) on all worksheets using UpdateSelectedValue, recalculates formulas, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string mainPath = "MainWorkbook.xlsx";

                // Verify main workbook exists
                if (!File.Exists(mainPath))
                {
                    Console.WriteLine($"Main workbook not found: {mainPath}");
                    return;
                }

                // Load the main workbook that contains linked shapes and external references
                Workbook mainWorkbook = new Workbook(mainPath);

                // Load external workbooks that serve as data sources for the links
                string[] externalPaths = { "ExternalData1.xlsx", "ExternalData2.xlsx" };
                Workbook[] externalWorkbooks = new Workbook[externalPaths.Length];

                for (int i = 0; i < externalPaths.Length; i++)
                {
                    if (!File.Exists(externalPaths[i]))
                    {
                        Console.WriteLine($"External workbook not found: {externalPaths[i]}");
                        return;
                    }

                    externalWorkbooks[i] = new Workbook(externalPaths[i]);
                }

                // Update all external links in the main workbook with the loaded data sources
                mainWorkbook.UpdateLinkedDataSource(externalWorkbooks);

                // Refresh linked shapes (e.g., dropdowns, list boxes) in every worksheet
                foreach (Worksheet sheet in mainWorkbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        try
                        {
                            shape.UpdateSelectedValue();
                        }
                        catch
                        {
                            // Ignore shapes that do not support UpdateSelectedValue
                        }
                    }
                }

                // Recalculate formulas to reflect the updated linked data
                mainWorkbook.CalculateFormula();

                // Save the updated workbook
                const string outputPath = "MainWorkbook_Updated.xlsx";
                mainWorkbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
