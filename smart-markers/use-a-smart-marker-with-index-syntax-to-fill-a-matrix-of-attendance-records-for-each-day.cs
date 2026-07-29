// Title: C# – Fill an Attendance Matrix with Aspose.Cells Smart Markers Using Array Index Syntax
// Description: Demonstrates how to create an Excel workbook, define a header row, insert smart markers that reference an employee name and indexed Attendance array elements, name the range "_CellsSmartMarkers" to repeat the template row for each employee, bind a List<Employee> as the data source, process the markers with WorkbookDesigner, and save the file.
// Keywords: Aspose.Cells | smart markers | array index syntax | C# | .NET | WorkbookDesigner | repeat rows | attendance matrix | Excel export | template row | data source binding
// Common Searches: Aspose.Cells smart markers array index example | C# repeat rows with _CellsSmartMarkers | How to fill Excel table using smart markers and arrays | Create attendance sheet with Aspose.Cells | Smart marker template for matrix data in .NET
// Developer Intent: Generate an Excel sheet where each employee’s daily attendance values are populated automatically via smart markers with index notation.
// Use Cases: Produce a daily attendance report that expands automatically for any number of employees. | Build a timesheet matrix where each column represents a day and rows are generated from a collection with an array property. | Create a scalable summary sheet that adapts to varying numbers of days or employees without manual column adjustments.
// AI Prompts: Show how to modify the smart‑marker template to support a dynamic number of day columns instead of hard‑coded indexes. | Explain how to bind a DataTable containing attendance columns to the smart markers using array‑index syntax. | Provide code that adds conditional formatting (e.g., highlight values < 8) to the attendance cells after processing the smart markers.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerMatrixDemo
{
    // Simple data class representing an employee and his/her attendance per day
    // Demonstrates how to create an Excel workbook, define a header row, insert smart markers that reference an employee name and indexed Attendance array elements, name the range "_CellsSmartMarkers" to repeat the template row for each employee, bind a List<Employee> as the data source, process the markers with WorkbookDesigner, and save the file.
    public class Employee
    {
        public string Name { get; set; } = null!;
        // Attendance array where each element corresponds to a day
        public int[] Attendance { get; set; } = null!;
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Create a new workbook
                Workbook wb = new Workbook();

                // 2. Access the first worksheet and its cells collection
                Worksheet ws = wb.Worksheets[0];
                Cells cells = ws.Cells;

                // 3. Set header row (Day names)
                cells["A1"].PutValue("Employee");
                cells["B1"].PutValue("Day 1");
                cells["C1"].PutValue("Day 2");
                cells["D1"].PutValue("Day 3");

                // 4. Insert smart markers in the template row (row 2)
                // Simple smart marker for employee name
                cells["A2"].PutValue("&=Employees.Name");
                // Index syntax to refer to array elements of Attendance
                cells["B2"].PutValue("&=Employees.Attendance[0]");
                cells["C2"].PutValue("&=Employees.Attendance[1]");
                cells["D2"].PutValue("&=Employees.Attendance[2]");

                // 5. Define the range that contains the smart markers and name it "_CellsSmartMarkers"
                // This enables the designer to repeat the row for each employee in the data source.
                Aspose.Cells.Range smRange = cells.CreateRange("A2:D2");
                smRange.Name = "_CellsSmartMarkers";

                // 6. Prepare sample data
                List<Employee> employees = new List<Employee>
                {
                    new Employee { Name = "Alice", Attendance = new int[] { 8, 9, 7 } },
                    new Employee { Name = "Bob",   Attendance = new int[] { 6, 8, 8 } },
                    new Employee { Name = "Carol", Attendance = new int[] { 9, 9, 10 } }
                };

                // 7. Set up the workbook designer
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = wb
                    // LineByLine is obsolete; range smart markers are used by default
                };
                designer.SetDataSource("Employees", employees);

                // 8. Process the smart markers
                designer.Process();

                // 9. Save the result
                string outputPath = "SmartMarkerMatrixDemo.xlsx";
                string fullPath = Path.GetFullPath(outputPath);
                string? outputDir = Path.GetDirectoryName(fullPath);

                // Ensure the output directory exists (handle possible null)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                wb.Save(fullPath);
                Console.WriteLine($"Workbook saved to: {fullPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
