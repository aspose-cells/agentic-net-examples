using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPerformanceLogging
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Stopwatch to measure each step
                Stopwatch sw = new Stopwatch();

                // 1. Load workbook (template)
                const string templatePath = "Template.xlsx"; // replace with actual template path
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                sw.Start();
                Workbook workbook = new Workbook(templatePath);
                sw.Stop();
                Console.WriteLine($"Load workbook: {sw.ElapsedMilliseconds} ms");
                sw.Reset();

                // 2. Prepare data source for smart markers
                sw.Start();
                var employees = new List<Employee>
                {
                    new Employee { Name = "John Doe", Age = 30, Department = "Sales" },
                    new Employee { Name = "Jane Smith", Age = 28, Department = "HR" }
                };
                sw.Stop();
                Console.WriteLine($"Prepare data source: {sw.ElapsedMilliseconds} ms");
                sw.Reset();

                // 3. Set up WorkbookDesigner and assign data source
                sw.Start();
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook,
                    LineByLine = false // using range smart markers
                };
                designer.SetDataSource("Employees", employees);
                sw.Stop();
                Console.WriteLine($"Configure designer & set data source: {sw.ElapsedMilliseconds} ms");
                sw.Reset();

                // 4. Process smart markers (localization step)
                sw.Start();
                designer.Process(); // processes all smart markers in the defined range
                sw.Stop();
                Console.WriteLine($"Process smart markers: {sw.ElapsedMilliseconds} ms");
                sw.Reset();

                // 5. Optional: calculate formulas with a custom monitor to log each cell calculation
                sw.Start();
                CalculationOptions calcOptions = new CalculationOptions
                {
                    CalculationMonitor = new PerformanceCalculationMonitor()
                };
                workbook.CalculateFormula(calcOptions);
                sw.Stop();
                Console.WriteLine($"Calculate formulas: {sw.ElapsedMilliseconds} ms");
                sw.Reset();

                // 6. Save the resulting workbook
                const string resultPath = "Result.xlsx"; // replace with desired output path
                sw.Start();
                workbook.Save(resultPath);
                sw.Stop();
                Console.WriteLine($"Save workbook: {sw.ElapsedMilliseconds} ms");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Sample data class used for smart markers
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string Department { get; set; }
    }

    // Custom calculation monitor to log each cell calculation (optional performance insight)
    public class PerformanceCalculationMonitor : AbstractCalculationMonitor
    {
        public override void BeforeCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Placeholder for per‑cell timing if needed
        }

        public override void AfterCalculate(int sheetIndex, int rowIndex, int columnIndex)
        {
            // Log after each cell is calculated (commented out to avoid excessive output)
            // Console.WriteLine($"Calculated Sheet{sheetIndex} Row{rowIndex} Col{columnIndex}");
        }
    }
}