using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SmartMarkerConditionalExample
{
    // Simple data class used as a data source
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Helper singleton to give the callback access to the workbook instance
    public sealed class WorkbookHolder
    {
        private static readonly Lazy<WorkbookHolder> _lazy = new(() => new WorkbookHolder());
        public static WorkbookHolder Instance => _lazy.Value;
        public Workbook Workbook { get; set; }
        private WorkbookHolder() { }
    }

    // Callback that removes rows where Age does not satisfy the condition (> _minAge)
    public class AgeFilterCallback : ISmartMarkerCallBack
    {
        private readonly int _minAge;
        public AgeFilterCallback(int minAge) => _minAge = minAge;

        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // We are only interested in the "Age" column
            if (!columnName.Equals("Age", StringComparison.OrdinalIgnoreCase))
                return;

            // Get the workbook/worksheet that is being processed
            Workbook wb = WorkbookHolder.Instance.Workbook;
            Worksheet ws = wb.Worksheets[sheetIndex];
            Cell ageCell = ws.Cells[rowIndex, colIndex];

            // Convert the cell value to int (smart marker already placed the value)
            if (int.TryParse(ageCell.Value?.ToString(), out int age))
            {
                // If the condition is not met, delete the entire row
                if (age <= _minAge)
                {
                    ws.Cells.DeleteRow(rowIndex);
                }
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers.
            // The template should have a range named "_CellsSmartMarkers"
            // and smart markers like &Persons.Name and &Persons.Age.
            Workbook wb = new Workbook("template.xlsx");

            // Store the workbook reference for the callback.
            WorkbookHolder.Instance.Workbook = wb;

            // Prepare sample data.
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John",  Age = 28 },
                new Person { Name = "Alice", Age = 35 },
                new Person { Name = "Bob",   Age = 42 }
            };

            // Create a WorkbookDesigner and bind the data source.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = wb,
                // Assign the callback that implements the conditional logic.
                CallBack = new AgeFilterCallback(minAge: 30)
            };
            designer.SetDataSource("Persons", persons);

            // Process the smart markers. The callback will remove rows where Age <= 30.
            designer.Process();

            // Save the populated workbook.
            wb.Save("output.xlsx");
        }
    }
}