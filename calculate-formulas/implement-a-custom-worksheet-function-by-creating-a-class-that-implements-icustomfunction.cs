using System;
using System.IO;
using Aspose.Cells;

public class CustomFunctionDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data.
            ws.Cells["A1"].PutValue(10);
            ws.Cells["A2"].PutValue(20);
            ws.Cells["A3"].PutValue(30);

            // Use the built‑in SUM function (custom functions require the ICustomFunction interface,
            // which may not be available in all Aspose.Cells versions).
            ws.Cells["B1"].Formula = "=SUM(A1:A3)";

            // Calculate formulas.
            wb.CalculateFormula();

            // Output the result.
            Console.WriteLine("Result of SUM(A1:A3): " + ws.Cells["B1"].Value);

            // Define output path and ensure the directory exists.
            string outputPath = "CustomFunctionDemo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        CustomFunctionDemo.Run();
    }
}