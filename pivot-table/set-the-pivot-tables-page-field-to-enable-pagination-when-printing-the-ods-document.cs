// Title: C# – Enable Pivot Table Page Field Pagination in ODS with Aspose.Cells
// Description: Loads an ODS workbook, finds the first worksheet that contains a pivot table, iterates over its page fields and calls ShowReportFilterPage to create a separate report‑filter page for each field, then saves the file. This prepares the pivot table for paginated printing.
// Keywords: Aspose.Cells C# ODS pagination | ShowReportFilterPage example | pivot table page fields printing | enable pivot pagination Aspose | ODS pivot table pagination
// Common Searches: Aspose.Cells enable pagination for pivot table page fields | ShowReportFilterPage C# ODS | print pivot table with separate filter pages | find first pivot table in workbook Aspose.Cells | how to paginate pivot table page fields
// Developer Intent: Add paginated report‑filter pages to each page field of a pivot table in an ODS workbook using Aspose.Cells for .NET.
// Use Cases: Generate a distinct filter page for every page field before printing an ODS pivot table. | Automatically apply pagination only when a workbook contains a pivot table. | Update an existing ODS document to support page‑field pagination without rebuilding the pivot table.
// AI Prompts: Write C# code that opens an ODS file, locates the first pivot table, and calls ShowReportFilterPage for each page field using Aspose.Cells. | Explain how ShowReportFilterPage affects the printed layout of a pivot table and why it creates separate filter pages. | Provide error‑handling best practices for missing files or absent pivot tables when enabling pagination in an ODS workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an ODS workbook, finds the first worksheet that contains a pivot table, iterates over its page fields and calls ShowReportFilterPage to create a separate report‑filter page for each field, then saves the file. This prepares the pivot table for paginated printing.
class EnablePivotPagination
{
    static void Main()
    {
        const string inputPath = "input.ods";
        const string outputPath = "output.ods";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing ODS workbook containing a pivot table
            Workbook workbook = new Workbook(inputPath);

            // Locate the first worksheet that contains at least one pivot table
            PivotTable pivotTable = null;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.PivotTables.Count > 0)
                {
                    pivotTable = ws.PivotTables[0];
                    break;
                }
            }

            if (pivotTable == null)
            {
                Console.WriteLine("No pivot tables were found in the workbook.");
                return;
            }

            // Enable pagination by generating a separate report filter page for each page field
            foreach (PivotField pageField in pivotTable.PageFields)
            {
                pivotTable.ShowReportFilterPage(pageField);
            }

            // Save the workbook back to ODS format
            workbook.Save(outputPath, SaveFormat.Ods);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
