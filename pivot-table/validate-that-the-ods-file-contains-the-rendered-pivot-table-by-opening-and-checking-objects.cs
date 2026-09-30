// Title: Validate that an ODS workbook contains a rendered pivot table with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that loads an ODS file, iterates all worksheets, and returns true when a pivot table has at least one row, column, or data field. | Create a C# method that checks an ODS workbook for any rendered pivot tables by examining each worksheet's PivotTableCollection via Aspose.Cells.
// Common Searches: how to programmatically confirm a pivot table exists in an ODS file using Aspose.Cells C# | C# Aspose.Cells example to detect rendered pivot tables in OpenDocument spreadsheets | verify ODS workbook contains pivot tables with row or column fields via Aspose.Cells | sample code for checking pivot table fields in ODS using Aspose.Cells for .NET | Aspose.Cells ODS pivot table validation tutorial
// Tags: Aspose.Cells ODS pivot table detection | C# validate rendered pivot table | PivotTableCollection iteration Aspose.Cells | check pivot fields ODS workbook | verify pivot table presence .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// Loads an ODS workbook with Aspose.Cells, iterates each worksheet's PivotTableCollection, and returns true if any pivot table has at least one row, column, or data field, indicating the pivot table is rendered.
class PivotTableValidator
{
    // Checks whether the specified ODS file contains at least one rendered pivot table.
    static bool ValidatePivotTable(string odsFilePath)
    {
        try
        {
            // Load the ODS workbook.
            Workbook workbook = new Workbook(odsFilePath);

            // Scan each worksheet for pivot tables.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                PivotTableCollection pivots = sheet.PivotTables;
                if (pivots != null && pivots.Count > 0)
                {
                    // Verify that the pivot table has fields (i.e., it is rendered).
                    foreach (PivotTable pt in pivots)
                    {
                        if (pt.RowFields.Count > 0 || pt.ColumnFields.Count > 0 || pt.DataFields.Count > 0)
                        {
                            return true; // Found a rendered pivot table.
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while validating pivot table: {ex.Message}");
        }

        return false; // No rendered pivot table detected or an error occurred.
    }

    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Usage: PivotTableValidator <path-to-ods-file>");
            return;
        }

        string odsPath = args[0];

        if (!File.Exists(odsPath))
        {
            Console.WriteLine($"File not found: {odsPath}");
            return;
        }

        bool hasPivot = ValidatePivotTable(odsPath);
        Console.WriteLine(hasPivot ? "Pivot table is present and rendered." : "No rendered pivot table found.");
    }
}
