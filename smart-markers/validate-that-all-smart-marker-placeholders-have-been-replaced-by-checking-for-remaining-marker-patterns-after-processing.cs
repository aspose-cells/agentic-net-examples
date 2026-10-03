// Title: C# code to confirm all Aspose.Cells smart marker placeholders are replaced after WorkbookDesigner processing
// AI Prompts: Write a C# routine that runs WorkbookDesigner.Process() on a workbook and then scans every cell with a regex to report any remaining ${...} or #=...# smart markers. | Create a method that iterates through all worksheets in an Aspose.Cells workbook, uses a pattern matcher to detect unreplaced smart marker syntax, and logs the sheet name and cell address for each occurrence.
// Common Searches: how to check for unreplaced smart markers in an Aspose.Cells workbook using C# | C# verify that ${...} placeholders are removed after processing with WorkbookDesigner | regex pattern to find #=...# smart markers in Excel after Aspose.Cells processing | detect leftover smart marker syntax in saved Excel file C# Aspose.Cells | post‑processing validation of smart markers in Excel with Aspose.Cells API
// Tags: smart marker replacement verification Aspose.Cells | regex detection of ${} placeholders in Excel | WorkbookDesigner post‑process validation C# | scan workbook cells for smart marker patterns | unreplaced smart marker detection C#

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook, processes its smart markers using WorkbookDesigner, saves the result, and then iterates through every worksheet and cell. A regular expression identifies any remaining ${...} or #=...# placeholders, logging their locations and reporting whether all smart markers were successfully replaced.
class Program
{
    static void Main()
    {
        // Load the workbook that contains smart markers
        Workbook workbook = new Workbook("input.xlsx");

        // Process smart markers (assumes data sources are already set on the designer)
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        // Example: designer.SetDataSource("Data", dataTable);
        designer.Process();

        // Save the processed workbook
        workbook.Save("output.xlsx");

        // Validate that no smart marker placeholders remain
        bool markersFound = false;
        // Regex matches common smart marker patterns like ${...} or #=...#
        Regex markerPattern = new Regex(@"\$\{.*?\}|\#\=.+?\#");

        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Type == CellValueType.IsString && markerPattern.IsMatch(cell.StringValue))
                {
                    markersFound = true;
                    Console.WriteLine($"Remaining marker found in sheet '{sheet.Name}' cell {cell.Name}: {cell.StringValue}");
                }
            }
        }

        if (!markersFound)
        {
            Console.WriteLine("All smart markers have been replaced successfully.");
        }
        else
        {
            Console.WriteLine("Some smart markers were not replaced.");
        }
    }
}
