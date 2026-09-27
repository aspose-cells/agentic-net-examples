// Title: Import a list of Employee objects into an Aspose.Cells worksheet and format the HireDate column as dd-MMM-yyyy using C#
// AI Prompts: Import an array of Employee objects into a worksheet with a header row using Aspose.Cells for .NET. | Create a style with a custom date pattern and apply it only to the HireDate column range. | Save the workbook as an .xlsx file after formatting the date column.
// Common Searches: how to import custom objects into Aspose.Cells worksheet with headers | apply custom date format dd-MMM-yyyy in Aspose.Cells C# | set number format for a specific column after ImportCustomObjects Aspose.Cells | format DateTime column in exported Excel using Aspose.Cells .NET
// Tags: populate worksheet from object collection Aspose.Cells | apply custom date style Aspose.Cells | format HireDate column Aspose.Cells | custom object list export to Excel .NET | style flag numberformat Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDateFormatExample
{
    // Define a custom object with a DateTime property
    // Shows how to import a collection of Employee objects into a worksheet with field names, apply a custom "dd-MMM-yyyy" date style to the HireDate column, and save the workbook as EmployeesWithFormattedDates.xlsx.
    public class Employee
    {
        public string Name { get; set; }
        public DateTime HireDate { get; set; }
        public double Salary { get; set; }

        public Employee(string name, DateTime hireDate, double salary)
        {
            Name = name;
            HireDate = hireDate;
            Salary = salary;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a list of custom objects
                List<Employee> employees = new List<Employee>
                {
                    new Employee("Alice", new DateTime(2020, 1, 15), 75000),
                    new Employee("Bob",   new DateTime(2019, 6, 30), 82000),
                    new Employee("Carol", new DateTime(2021, 11, 5), 68000)
                };

                // Import the list into the worksheet starting at cell A1
                // The first row will contain property names (Name, HireDate, Salary)
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    IsFieldNameShown = true   // include header row
                };
                sheet.Cells.ImportCustomObjects(employees.ToArray(), 0, 0, importOptions);

                // Apply custom number format to the HireDate column (column B, index 1)
                // The format string "dd-MMM-yyyy" will display dates like "15-Jan-2020"
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Custom = "dd-MMM-yyyy";

                // Create a style flag to apply only the number format
                StyleFlag flag = new StyleFlag
                {
                    NumberFormat = true
                };

                // Determine the range that contains the date values (excluding header)
                int firstDataRow = 1; // zero‑based index; row 1 is the second row (after header)
                Aspose.Cells.Range dateRange = sheet.Cells.CreateRange(firstDataRow, 1, employees.Count, 1); // column B

                // Apply the style to the range
                dateRange.ApplyStyle(dateStyle, flag);

                // Save the workbook to a file
                workbook.Save("EmployeesWithFormattedDates.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
