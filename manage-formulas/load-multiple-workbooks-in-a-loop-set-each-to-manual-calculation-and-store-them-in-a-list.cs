// Title: C# – Load Multiple Excel Workbooks, Set Manual Calculation Mode, and Store in a List with Aspose.Cells
// Description: C# example that iterates over an array of file paths, verifies each file, loads it into an Aspose.Cells Workbook, switches the FormulaSettings.CalculationMode to Manual, adds the workbook to a List, and returns the collection while handling missing files and load errors.
// Keywords: Aspose.Cells | C# | load multiple workbooks | manual calculation mode | FormulaSettings | batch Excel processing | Workbook list | disable auto calculation | Excel file loop | error handling
// Common Searches: Aspose.Cells load several Excel files into a list | set calculation mode manual for each workbook C# | batch load workbooks without auto recalculation Aspose.Cells | C# loop to open multiple Excel workbooks Aspose.Cells | prevent formula evaluation when loading workbooks Aspose.Cells
// Developer Intent: Open a series of Excel files, turn off automatic formula evaluation, and keep the Workbook objects in a collection for further processing.
// Use Cases: Pre‑process a batch of reports without triggering costly recalculations, then run custom calculations on demand. | Merge or compare data from several workbooks while preserving performance by disabling auto‑calc. | Load workbooks for background or server‑side tasks where formulas should be evaluated only when explicitly required.
// AI Prompts: Generate C# code that accepts an array of Excel file paths, loads each into an Aspose.Cells Workbook, sets FormulaSettings.CalculationMode to Manual, and returns a List<Workbook> with robust error handling. | Show how to iterate over the List<Workbook> returned by LoadWorkbooks and export every worksheet to CSV using Aspose.Cells. | Modify the LoadWorkbooks method to accept a collection of Stream objects instead of file paths while keeping manual calculation mode enabled.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// C# example that iterates over an array of file paths, verifies each file, loads it into an Aspose.Cells Workbook, switches the FormulaSettings.CalculationMode to Manual, adds the workbook to a List, and returns the collection while handling missing files and load errors.
public class WorkbookLoader
{
    /// <param name="filePaths">Array of full file paths to the Excel files.</param>
    /// <returns>List containing the loaded Workbook objects.</returns>
    public List<Workbook> LoadWorkbooks(string[] filePaths)
    {
        var workbooks = new List<Workbook>();

        foreach (string path in filePaths)
        {
            try
            {
                // Verify that the file exists before attempting to load it
                if (!File.Exists(path))
                {
                    Console.WriteLine($"Warning: File not found – skipping '{path}'.");
                    continue;
                }

                // Load the workbook from the file
                var wb = new Workbook(path);

                // Set calculation mode to Manual
                wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                workbooks.Add(wb);
            }
            catch (Exception ex)
            {
                // Log the exception and continue processing other files
                Console.WriteLine($"Error loading workbook '{path}': {ex.Message}");
            }
        }

        return workbooks;
    }
}

// Example usage
class Program
{
    static void Main()
    {
        // Define the Excel files to load
        string[] files = new string[]
        {
            @"C:\Data\Report1.xlsx",
            @"C:\Data\Report2.xlsx",
            @"C:\Data\Report3.xlsx"
        };

        var loader = new WorkbookLoader();
        List<Workbook> loadedWorkbooks = loader.LoadWorkbooks(files);

        Console.WriteLine($"Loaded {loadedWorkbooks.Count} workbooks with Manual calculation mode.");
    }
}
