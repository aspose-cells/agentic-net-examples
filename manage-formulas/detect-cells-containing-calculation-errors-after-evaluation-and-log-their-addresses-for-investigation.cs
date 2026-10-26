// Title: Find and log Excel cells with calculation errors after recalculating formulas using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to list cells that return #ERROR after workbook.CalculateFormula | Aspose.Cells .NET detect cells with formula errors and get their addresses | Retrieve addresses of error cells in an Excel workbook using Aspose.Cells C# | Log calculation error cells after recalculating formulas with Aspose.Cells in a .NET application | Find cells with CellValueType.IsError in Aspose.Cells after evaluating formulas
// Tags: recalculate formulas Aspose.Cells | detect error cells Aspose.Cells | CellValueType.IsError C# | list error cell addresses Excel | load password protected workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook (prompting for a password if required), recalculates all formulas with Workbook.CalculateFormula, scans each worksheet for cells whose type is CellValueType.IsError, collects their sheet!A1 addresses, writes the list to the console, and saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook = null;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            // If loading fails, assume the file may be password‑protected
            Console.WriteLine("Failed to load workbook: " + ex.Message);
            Console.Write("If the file is password protected, please enter the password (or press Enter to abort): ");
            string password = Console.ReadLine();

            if (string.IsNullOrEmpty(password))
            {
                Console.WriteLine("No password provided. Exiting.");
                return;
            }

            try
            {
                // Load with the supplied password
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx) { Password = password };
                workbook = new Workbook(inputPath, loadOptions);
            }
            catch (Exception innerEx)
            {
                Console.WriteLine("Unable to load workbook with the provided password: " + innerEx.Message);
                return;
            }
        }

        try
        {
            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            List<string> errorCellAddresses = new List<string>();

            // Scan each worksheet for cells that contain errors
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        if (cell.Type == CellValueType.IsError)
                        {
                            string address = $"{sheet.Name}!{cell.Name}";
                            errorCellAddresses.Add(address);
                        }
                    }
                }
            }

            // Output the addresses of error cells
            Console.WriteLine("Cells containing calculation errors:");
            foreach (string address in errorCellAddresses)
            {
                Console.WriteLine(address);
            }

            // Save the workbook after processing
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during processing: " + ex.Message);
        }
    }
}
