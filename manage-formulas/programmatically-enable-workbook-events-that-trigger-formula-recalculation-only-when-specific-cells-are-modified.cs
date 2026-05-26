using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom monitor to log calculation events (optional)
    public class LoggingCalculationMonitor : AbstractCalculationMonitor
    {
        public override void AfterCalculate(int sheetIndex, int rowIndex, int colIndex)
        {
            Console.WriteLine($"Calculated cell: Sheet{sheetIndex} {CellsHelper.CellIndexToName(rowIndex, colIndex)}");
        }
    }

    public class WorkbookEventDrivenRecalcDemo
    {
        public static void Run()
        {
            // 1. Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // 2. Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // 3. Set up some initial data and formulas
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].Formula = "=A1+A2";   // depends on A1 and A2
            cells["C1"].Formula = "=B1*2";    // depends on B1

            // 4. Configure formula calculation to Manual mode
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // 5. Define the set of cells whose modification should trigger recalculation
            HashSet<string> triggerCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "A1",   // when A1 changes, recalc
                "A2"    // when A2 changes, recalc
                // add more cell addresses as needed
            };

            // 6. Prepare calculation options with a monitor (optional, for logging)
            CalculationOptions calcOptions = new CalculationOptions
            {
                CalculationMonitor = new LoggingCalculationMonitor()
            };

            // 7. Helper method to modify a cell and conditionally recalculate
            void ModifyCell(string address, object value)
            {
                Console.WriteLine($"\nModifying {address} to {value}");
                cells[address].PutValue(value);

                // If the modified cell is in the trigger list, recalculate formulas
                if (triggerCells.Contains(address))
                {
                    Console.WriteLine($"Trigger cell changed. Recalculating formulas...");
                    workbook.CalculateFormula(calcOptions);
                }
                else
                {
                    Console.WriteLine($"Cell {address} is not a trigger cell. Skipping recalculation.");
                }
            }

            // 8. Perform modifications
            ModifyCell("A1", 30);   // triggers recalculation
            ModifyCell("B2", 5);    // does NOT trigger recalculation
            ModifyCell("A2", 40);   // triggers recalculation

            // 9. Save the workbook (lifecycle rule: save)
            workbook.Save("EventDrivenRecalcOutput.xlsx", SaveFormat.Xlsx);

            // 10. Display final values for verification
            Console.WriteLine("\nFinal cell values after conditional recalculations:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");
            Console.WriteLine($"B1 (formula) = {cells["B1"].Value}");
            Console.WriteLine($"C1 (formula) = {cells["C1"].Value}");
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            WorkbookEventDrivenRecalcDemo.Run();
        }
    }
}