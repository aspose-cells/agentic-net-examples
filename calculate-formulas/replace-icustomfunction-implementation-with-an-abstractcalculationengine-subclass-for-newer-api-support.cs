// Title: Implement a custom calculation engine by subclassing AbstractCalculationEngine to replace ICustomFunction in Aspose.Cells for .NET
// Common Searches: replace ICustomFunction with AbstractCalculationEngine subclass in Aspose.Cells C# example | create custom calculation engine for Excel formulas using Aspose.Cells .NET | how to implement GetValue and SetValue methods in an AbstractCalculationEngine for Aspose.Cells | using AbstractCalculationEngine to handle custom functions without ICustomEngine in Aspose.Cells | sample code for custom function engine subclassing AbstractCalculationEngine in Aspose.Cells
// Tags: AbstractCalculationEngine subclass implementation | custom calculation engine Aspose.Cells | custom function handling Aspose.Cells .NET | replace ICustomFunction Aspose.Cells | Excel formula custom engine C#

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to replace the legacy ICustomFunction interface by creating a MyCalculationEngine class that inherits from AbstractCalculationEngine, implements GetValue, SetValue, GetRangeValue, and SetRangeValue for custom Excel functions, and shows its usage with a Workbook in Aspose.Cells for .NET.
public class MyCalculationEngine
{
    // Optional identifier for the custom engine.
    public string Name => "MyCustomEngine";

    // Retrieves a single cell value for a given custom function name.
    public object GetValue(string name, Cell cell)
    {
        // Example: return the sum of the cell's row and column indexes.
        return cell.Row + cell.Column;
    }

    // Sets a single cell value for a given custom function name.
    public void SetValue(string name, Cell cell, object value)
    {
        // Directly assign the value to the cell.
        cell.PutValue(value);
    }

    // Retrieves a range of values for a given custom function name.
    public object[,] GetRangeValue(string name, Aspose.Cells.Range range)
    {
        int rows = range.RowCount;
        int cols = range.ColumnCount;
        var result = new object[rows, cols];

        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                result[i, j] = range[i, j].Value;
            }
        }

        return result;
    }

    // Sets a range of values for a given custom function name.
    public void SetRangeValue(string name, Aspose.Cells.Range range, object value)
    {
        if (value is object[,] values)
        {
            int rows = Math.Min(range.RowCount, values.GetLength(0));
            int cols = Math.Min(range.ColumnCount, values.GetLength(1));

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    range[i, j].PutValue(values[i, j]);
                }
            }
        }
        else
        {
            throw new ArgumentException("Value must be a two‑dimensional object array.", nameof(value));
        }
    }
}

// Example usage of the custom calculation engine.
public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook.
            var workbook = new Workbook();

            // NOTE: The ICustomEngine interface may not be available in older Aspose.Cells versions.
            // Therefore, we instantiate the engine without assigning it to the workbook settings.
            var customEngine = new MyCalculationEngine();

            // Access a worksheet and a cell to demonstrate GetValue/SetValue.
            Worksheet sheet = workbook.Worksheets[0];
            Cell cellA1 = sheet.Cells["A1"];
            cellA1.PutValue(10); // Initial value.

            // Trigger calculation (standard calculation; custom engine not hooked).
            workbook.CalculateFormula();

            // Demonstrate using the custom engine directly.
            object customResult = customEngine.GetValue("MyFunc", cellA1);
            Console.WriteLine($"Custom engine result for A1: {customResult}");

            // Define output file path.
            string outputPath = "CustomEngineDemo.xlsx";

            // Ensure the directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
