using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class LoadWorkbooksManualCalc
    {
        /// <summary>
        /// Loads each workbook from the provided file paths, sets its calculation mode to Manual,
        /// and stores the workbook instances in a list.
        /// </summary>
        /// <param name="filePaths">Array of workbook file paths to load.</param>
        /// <returns>List containing the loaded Workbook objects.</returns>
        public static List<Workbook> LoadWorkbooks(string[] filePaths)
        {
            // List to hold the loaded workbooks
            List<Workbook> workbooks = new List<Workbook>();

            // Iterate over each file path
            foreach (string path in filePaths)
            {
                // Load the workbook using the string constructor (loads from file)
                Workbook wb = new Workbook(path);

                // Set calculation mode to Manual
                wb.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

                // Add the workbook to the collection
                workbooks.Add(wb);
            }

            return workbooks;
        }

        // Example usage
        public static void Main()
        {
            // Example file paths (replace with actual paths)
            string[] files = new string[]
            {
                "Workbook1.xlsx",
                "Workbook2.xlsx",
                "Workbook3.xlsx"
            };

            // Load workbooks with manual calculation mode
            List<Workbook> loadedWorkbooks = LoadWorkbooks(files);

            // Verify that the calculation mode is set to Manual
            foreach (Workbook wb in loadedWorkbooks)
            {
                Console.WriteLine($"Workbook '{wb.FileName}' calculation mode: {wb.Settings.FormulaSettings.CalculationMode}");
            }

            // Dispose workbooks when done (optional but recommended)
            foreach (Workbook wb in loadedWorkbooks)
            {
                wb.Dispose();
            }
        }
    }
}