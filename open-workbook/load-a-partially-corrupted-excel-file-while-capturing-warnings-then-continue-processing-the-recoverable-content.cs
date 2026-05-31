using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom warning callback that records all warnings
    public class CustomWarningCallback : IWarningCallback
    {
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo warningInfo)
        {
            // Store warning and also write to console
            Warnings.Add(warningInfo);
            Console.WriteLine($"Warning: {warningInfo.Type} - {warningInfo.Description}");
        }
    }

    public class LoadCorruptedWorkbookDemo
    {
        public static void Run()
        {
            // Path to the partially corrupted Excel file
            string filePath = "corrupted.xlsx";

            // Create load options and assign the custom warning callback
            LoadOptions loadOptions = new LoadOptions();
            CustomWarningCallback warningCallback = new CustomWarningCallback();
            loadOptions.WarningCallback = warningCallback;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Indicate that the workbook was opened in repair mode (optional but aligns with the scenario)
            workbook.Settings.RepairLoad = true;

            // Process recoverable content – example: print first few cells of each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");
                for (int row = 0; row < 5; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        Console.Write($"{sheet.Cells[row, col].StringValue}\t");
                    }
                    Console.WriteLine();
                }
            }

            // Optionally save the recovered workbook to a new file
            workbook.Save("recovered_output.xlsx");
        }
    }

    class Program
    {
        static void Main()
        {
            LoadCorruptedWorkbookDemo.Run();
        }
    }
}