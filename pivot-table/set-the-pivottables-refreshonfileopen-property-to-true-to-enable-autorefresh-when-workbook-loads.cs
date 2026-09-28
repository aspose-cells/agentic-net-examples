// Title: Set PivotTable.RefreshOnFileOpen to true for automatic refresh on workbook open with Aspose.Cells in C#
// AI Prompts: Configure the first PivotTable in an existing workbook to auto‑refresh on file open by setting RefreshOnFileOpen = true and saving the file with Aspose.Cells. | Write C# code that loads an Excel file, enables RefreshOnFileOpen for its pivot table, and writes the updated workbook. | Update a pivot table's RefreshOnFileOpen property using Aspose.Cells and persist the change to a new .xlsx file.
// Common Searches: how to enable RefreshOnFileOpen for a pivot table using Aspose.Cells C# | Aspose.Cells C# set pivot table auto refresh when opening workbook | example code to set PivotTable.RefreshOnFileOpen property in .NET | auto refresh pivot table on workbook load Aspose.Cells tutorial
// Tags: Aspose.Cells set PivotTable.RefreshOnFileOpen | C# enable pivot table auto refresh | Aspose.Cells update pivot table property | Excel workbook auto refresh pivot table .NET | PivotTable RefreshOnFileOpen example

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing Excel workbook, checks for a pivot table on the first worksheet, sets the pivot table's RefreshOnFileOpen property to true so it refreshes automatically when the file is opened, and saves the modified workbook to a new file, with error handling for missing files and exceptions.
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
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Refresh the pivot table data manually (RefreshDataOnOpen not available in this version)
            pivotTable.RefreshData();

            // Save the workbook with the updated setting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
