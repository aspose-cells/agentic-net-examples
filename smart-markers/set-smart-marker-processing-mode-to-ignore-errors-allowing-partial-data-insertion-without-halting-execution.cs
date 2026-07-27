using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerIgnoreErrorsDemo
{
    // Demonstrates how to process smart markers while ignoring errors.
    // Unrecognized or problematic smart markers are preserved instead of causing an exception.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add smart markers.
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Smart markers that reference a data source named "Employees".
            // The second marker refers to a non‑existent column "Salary" to simulate an error.
            sheet.Cells["A1"].PutValue("&=Employees.Name");
            sheet.Cells["B1"].PutValue("&=Employees.Salary"); // will cause an error

            // Define the range that contains smart markers.
            // When LineByLine is false, the designer processes only this named range.
            sheet.Cells.CreateRange("A1:B1").Name = "_CellsSmartMarkers";

            // -------------------------------------------------
            // 2. Prepare a data source with missing "Salary" column.
            // -------------------------------------------------
            var employees = new List<Employee>
            {
                new Employee { Name = "John Doe" },
                new Employee { Name = "Jane Smith" }
            };

            // -------------------------------------------------
            // 3. Configure the WorkbookDesigner.
            // -------------------------------------------------
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                // Setting LineByLine to false tells the designer to use the named range.
                LineByLine = false
            };

            // Register the data source.
            designer.SetDataSource("Employees", employees);

            // -------------------------------------------------
            // 4. Process smart markers while preserving unrecognized markers.
            //    The boolean parameter 'true' means: preserve markers that cannot be resolved.
            //    This effectively ignores errors and allows partial data insertion.
            // -------------------------------------------------
            designer.Process(isPreserved: true);

            // -------------------------------------------------
            // 5. Save the resulting workbook.
            // -------------------------------------------------
            workbook.Save("SmartMarkers_IgnoredErrors.xlsx");

            Console.WriteLine("Workbook saved. Unresolved smart markers were preserved.");
        }

        // Simple data class used as the data source.
        public class Employee
        {
            public string Name { get; set; }
            // Note: No Salary property – this will trigger an error that gets ignored.
        }
    }
}