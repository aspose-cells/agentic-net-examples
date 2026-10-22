// Title: Refresh a specific PivotTable in an existing Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, find a PivotTable by its name, invoke RefreshData, and write the updated file using Aspose.Cells in C#. | Update a PivotTable after source data changes and configure it to refresh automatically when the workbook is opened, using Aspose.Cells for .NET. | Implement comprehensive error handling for missing files, worksheets, or PivotTables while programmatically refreshing a PivotTable in a C# application.
// Common Searches: Aspose.Cells C# how to refresh a named PivotTable after changing source data | programmatically refresh a specific pivot table in a .NET workbook | set pivot table to auto‑refresh on opening with Aspose.Cells for .NET
// Tags: Aspose.Cells pivot table refresh operation | C# refresh specific pivot table | enable pivot auto refresh on workbook open Aspose | handle missing worksheet or pivot table Aspose.Cells | save workbook after pivot refresh .NET

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel file, validates the target worksheet and PivotTable by name, calls RefreshData on the PivotTable to incorporate any source data modifications, optionally enables automatic refresh on workbook opening, saves the workbook to a new file, and includes error handling for missing files, worksheets, or PivotTables.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook that contains the PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the worksheet where the PivotTable resides (adjust the name as needed)
            Worksheet sheet = workbook.Worksheets["Sheet1"];
            if (sheet == null)
            {
                Console.WriteLine("Error: Worksheet \"Sheet1\" not found in the workbook.");
                return;
            }

            // Retrieve the specific PivotTable by its name (adjust the name as needed)
            PivotTable pivot = sheet.PivotTables["PivotTable1"];
            if (pivot == null)
            {
                Console.WriteLine("Error: PivotTable \"PivotTable1\" not found on the worksheet.");
                return;
            }

            // Refresh the PivotTable so it reflects any changes made to its data source
            pivot.RefreshData();

            // Optional: ensure the PivotTable refreshes automatically when the workbook is opened
            // pivot.RefreshDataOnOpening = true;

            // Save the workbook with the refreshed PivotTable
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
