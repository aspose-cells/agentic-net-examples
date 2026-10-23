// Title: Import a 2‑D object array and a list of custom objects into the same worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells import 2D array and custom object list into one worksheet | How to combine ImportArray and ImportCustomObjects in Aspose.Cells .NET | Populate Excel sheet with primitive array data and Employee objects using Aspose.Cells | Aspose.Cells import heterogeneous data sources into a single worksheet C# example | Saving mixed array and object data to Excel with Aspose.Cells in .NET
// Tags: object[,] ImportArray usage Aspose.Cells | ImportCustomObjects with field name header C# | mixed primitive and custom object import workflow | Excel worksheet population from multiple .NET data collections | save combined data workbook Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Custom object definition for ImportCustomObjects
    // The example creates a new Workbook, uses Cells.ImportArray to write a 2‑D object array row by row, then employs Cells.ImportCustomObjects with ImportTableOptions to add a List<Employee> (including a header row) starting at cell A6, and finally saves the combined worksheet as CombinedData.xlsx.
    public class Employee
    {
        public int EmpId { get; set; }
        public string Dept { get; set; } = string.Empty;   // Initialize to avoid nullable warning
        public double Salary { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // {CreateWorkbook}
                Workbook workbook = new Workbook();   // Create a new workbook

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------- ImportArray --------------------
                // Heterogeneous primitive data (strings, ints, doubles)
                object[,] arrayData = new object[,]
                {
                    { "ID", "Name", "Score" },
                    { 1, "Alice", 85 },
                    { 2, "Bob", 92 },
                    { 3, "Charlie", 78 }
                };

                // Import the 2‑D array row by row, converting each row to string[]
                int rows = arrayData.GetLength(0);
                int cols = arrayData.GetLength(1);
                for (int r = 0; r < rows; r++)
                {
                    string[] row = new string[cols];
                    for (int c = 0; c < cols; c++)
                    {
                        row[c] = arrayData[r, c]?.ToString() ?? string.Empty;
                    }
                    sheet.Cells.ImportArray(row, r, 0, false);
                }

                // -------------------- ImportCustomObjects --------------------
                // Create a list of custom objects to be imported
                List<Employee> employees = new List<Employee>()
                {
                    new Employee(){ EmpId = 101, Dept = "HR", Salary = 50000 },
                    new Employee(){ EmpId = 102, Dept = "IT", Salary = 75000 },
                    new Employee(){ EmpId = 103, Dept = "Finance", Salary = 62000 }
                };

                // Define import options: treat the first row as a header (property names)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true
                };

                // Import the custom objects starting at cell A6 (row 5, column 0)
                sheet.Cells.ImportCustomObjects(employees, 5, 0, importOptions);

                // {SaveWorkbook}
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CombinedData.xlsx");

                // Ensure the directory exists before saving
                string? directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
