using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Extract data from specific cells
        string name = sheet.Cells["A1"].StringValue;

        // Safely get integer value from B1
        int age = 0;
        Cell ageCell = sheet.Cells["B1"];
        if (ageCell.Type == CellValueType.IsNumeric)
        {
            age = ageCell.IntValue;
        }
        else
        {
            int.TryParse(ageCell.StringValue, out age);
        }

        // Safely get double value from C1
        double salary = 0.0;
        Cell salaryCell = sheet.Cells["C1"];
        if (salaryCell.Type == CellValueType.IsNumeric)
        {
            salary = salaryCell.DoubleValue;
        }
        else
        {
            double.TryParse(salaryCell.StringValue, out salary);
        }

        // Create an anonymous object populated with the extracted data
        var employee = new
        {
            Name = name,
            Age = age,
            Salary = salary
        };

        // Example usage of the anonymous object
        Console.WriteLine($"Name: {employee.Name}, Age: {employee.Age}, Salary: {employee.Salary}");

        // Save the workbook (optional if no changes were made)
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}