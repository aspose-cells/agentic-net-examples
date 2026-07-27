// Title: Aspose.Cells .NET: Delete CheckBox Controls Linked to Zero‑Valued Cells in All Worksheets
// Description: Loads an Excel workbook, scans every worksheet for CheckBox form controls, checks the value of each control's LinkedCell, and removes the CheckBox when the linked cell contains the numeric value 0. The updated workbook is then saved.
// Keywords: Aspose.Cells delete checkboxes | remove zero linked checkbox .NET | iterate worksheets Aspose.Cells | Excel form control removal | C# Aspose.Cells shape manipulation | check box linked cell value | Aspose.Cells API example
// Common Searches: how to remove checkboxes linked to zero in Excel using Aspose.Cells | Aspose.Cells delete form controls based on cell value | C# iterate all worksheets and delete specific shapes | remove zero‑linked check boxes programmatically | Aspose.Cells example for cleaning up check box controls
// Developer Intent: Programmatically delete every CheckBox whose LinkedCell contains the numeric value zero across all worksheets in an Excel file.
// Use Cases: Clean a template before distribution by stripping unchecked (zero‑linked) check boxes. | Generate a final report without visual clutter from disabled options. | Automate workbook sanitization in data pipelines where zero‑linked controls are irrelevant.
// AI Prompts: Write C# code with Aspose.Cells that removes all check boxes linked to cells equal to 0 in every worksheet. | Show how to log the addresses of deleted check boxes while processing a workbook with Aspose.Cells. | Extend the example to also delete radio buttons whose linked cells contain zero.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Loads an Excel workbook, scans every worksheet for CheckBox form controls, checks the value of each control's LinkedCell, and removes the CheckBox when the linked cell contains the numeric value 0. The updated workbook is then saved.
    public class DeleteZeroLinkedCheckBoxes
    {
        // Entry point required for console application
        public static void Main(string[] args)
        {
            try
            {
                // Determine input and output file paths (allow overrides via command‑line)
                string inputPath = args.Length > 0 ? args[0] : "input.xlsx";
                string outputPath = args.Length > 1 ? args[1] : "output.xlsx";

                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                Run(inputPath, outputPath);
                Console.WriteLine($"Processing completed. Output saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Core logic moved to a method that accepts file paths
        public static void Run(string inputFile, string outputFile)
        {
            // Load the workbook from the specified input file
            Workbook workbook = new Workbook(inputFile);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of check boxes on the current worksheet
                CheckBoxCollection checkBoxes = sheet.CheckBoxes;

                // Iterate backwards because we will be removing items
                for (int i = checkBoxes.Count - 1; i >= 0; i--)
                {
                    CheckBox checkBox = checkBoxes[i];

                    // Get the cell address linked to the check box
                    string linkedCell = checkBox.LinkedCell;

                    // Proceed only if the check box is linked to a cell
                    if (!string.IsNullOrEmpty(linkedCell))
                    {
                        // Retrieve the cell object
                        Cell cell = sheet.Cells[linkedCell];

                        // Ensure the cell has a value that can be interpreted as a number
                        if (cell != null && cell.Value != null)
                        {
                            // Try to parse the cell value to a double
                            if (double.TryParse(cell.Value.ToString(), out double numericValue))
                            {
                                // If the linked cell contains zero, remove the check box
                                if (numericValue == 0)
                                {
                                    checkBoxes.RemoveAt(i);
                                }
                            }
                        }
                    }
                }
            }

            // Save the modified workbook to the specified output file
            workbook.Save(outputFile, SaveFormat.Xlsx);
        }
    }
}
