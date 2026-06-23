using System;
using Aspose.Cells;

namespace AsposeCellsDiagnostics
{
    public class LogFrozenState
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();

                // Add sample worksheets and freeze panes for demonstration
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "FirstSheet";
                // Freeze panes at cell C3 with 2 frozen rows and 2 frozen columns
                sheet1.FreezePanes(2, 2, 2, 2);

                Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
                // No freeze on this sheet

                Worksheet sheet3 = workbook.Worksheets.Add("ThirdSheet");
                // Freeze panes at cell B2 with 1 frozen row and 1 frozen column
                sheet3.FreezePanes(1, 1, 1, 1);

                // Iterate through all worksheets and log their frozen state
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    int row, column, frozenRows, frozenColumns;
                    bool hasFreeze = ws.GetFreezedPanes(out row, out column, out frozenRows, out frozenColumns);

                    Console.WriteLine($"Worksheet '{ws.Name}': Frozen = {hasFreeze}");
                    if (hasFreeze)
                    {
                        Console.WriteLine($"  Freeze Position - Row: {row}, Column: {column}");
                        Console.WriteLine($"  Frozen Rows: {frozenRows}, Frozen Columns: {frozenColumns}");
                    }
                }

                // Save the workbook (lifecycle rule: save)
                workbook.Save("FrozenStateDiagnostics.xlsx");
                Console.WriteLine("Workbook saved as 'FrozenStateDiagnostics.xlsx'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            LogFrozenState.Run();
        }
    }
}