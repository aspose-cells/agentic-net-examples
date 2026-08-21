// Title: C# – Load Multiple Excel Workbooks, Set Manual Calculation Mode, and Store in a List with Aspose.Cells
// Description: A concise example that loops through an array of Excel file paths, creates an Aspose.Cells Workbook for each file, switches the FormulaSettings.CalculationMode to Manual, and adds the workbook objects to a List<Workbook>. Ideal for batch processing without triggering automatic recalculation.
// Keywords: Aspose.Cells load multiple workbooks | C# manual calculation mode | FormulaSettings.CalculationMode Manual | batch workbook processing .NET | store Workbook objects in List | Aspose.Cells example GitHub | disable auto calculation Aspose.Cells | Excel bulk load C#
// Common Searches: how to load several Excel files with Aspose.Cells and set manual calculation | batch processing workbooks in C# using Aspose.Cells | store multiple Aspose.Cells Workbook instances in a collection | disable automatic formula calculation when opening Excel files Aspose.Cells
// Developer Intent: Load each workbook, set its formula calculation to manual, and keep the instances in a List for later processing.
// Use Cases: Prepare a set of workbooks for bulk data updates without triggering recalculation after each file is opened. | Open many workbooks, modify cell values, and run a single manual calculation pass when needed. | Collect workbooks in a list to pass them to a reporting engine or export routine after configuring manual calculation.
// AI Prompts: Generate C# code that iterates over an array of Excel file paths, loads each file with Aspose.Cells, sets FormulaSettings.CalculationMode to Manual, and adds the Workbook to a List<Workbook>. | Show how to recalculate formulas manually for all workbooks stored in a List after making data changes using Aspose.Cells. | Explain the steps to prevent automatic formula evaluation when loading multiple Excel workbooks in a loop with Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsMultipleWorkbooks
{
    // A concise example that loops through an array of Excel file paths, creates an Aspose.Cells Workbook for each file, switches the FormulaSettings.CalculationMode to Manual, and adds the workbook objects to a List<Workbook>. Ideal for batch processing without triggering automatic recalculation.
    class Program
    {
        static void Main()
        {
            // List of file paths to load
            string[] filePaths = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
                // Add more paths as needed
            };

            // List to hold loaded workbooks
            List<Workbook> workbooks = new List<Workbook>();

            // Loop through each file, load it, set manual calculation, and store in the list
            foreach (string path in filePaths)
            {
                // Load workbook using the string constructor (load rule)
                Workbook wb = new Workbook(path);

                // Set calculation mode to Manual (FormulaSettings.CalculationMode)
                wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // Add the workbook to the collection
                workbooks.Add(wb);
            }

            // At this point, 'workbooks' contains all loaded workbooks with manual calculation mode.
            Console.WriteLine($"Loaded {workbooks.Count} workbooks with Manual calculation mode.");
        }
    }
}
