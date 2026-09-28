// Title: Bind an array of anonymous .NET objects to a smart marker and highlight salaries over 6000 with conditional formatting using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates an anonymous object array for employees, assigns it to a WorkbookDesigner data source named "Employees", and processes the smart markers in a new workbook. | Add C# statements that define a conditional formatting rule on the Salary column to set a yellow background when the cell value is greater than 6000. | Write the code to save the processed workbook as an XLSX file named "AnonymousDataSmartMarkers.xlsx" and handle any exceptions.
// Common Searches: asp.net bind anonymous list to smart markers Aspose.Cells example | how to apply conditional formatting to smart marker generated cells in C# | highlight salary column values greater than 6000 using Aspose.Cells conditional formatting | using WorkbookDesigner with anonymous objects array for smart markers in .NET
// Tags: anonymous objects data source smart markers | conditional formatting cell value greater than threshold Aspose.Cells | WorkbookDesigner set data source array C# | highlight high salary cells yellow Aspose.Cells | smart markers with conditional formatting .NET

using System;
using System.Drawing;
using Aspose.Cells;

// The sample creates a workbook, inserts smart markers for employee fields, builds an array of anonymous objects, binds it to the "Employees" marker via WorkbookDesigner, processes the markers, adds a conditional formatting rule that colors Salary cells above 6000 yellow, and saves the result as AnonymousDataSmartMarkers.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // 2. Place smart markers in the template.
            //    The marker name "Employees" will be bound to the data source later.
            sheet.Cells["A1"].PutValue("&=Employees.Name");
            sheet.Cells["B1"].PutValue("&=Employees.Age");
            sheet.Cells["C1"].PutValue("&=Employees.Salary");

            // 3. Build a collection of anonymous objects.
            var employees = new[]
            {
                new { Name = "John", Age = 30, Salary = 5000 },
                new { Name = "Jane", Age = 45, Salary = 8000 },
                new { Name = "Bob",  Age = 28, Salary = 4000 }
            };

            // 4. Create a WorkbookDesigner, assign the workbook and set the data source.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource("Employees", employees);

            // 5. Process the smart markers – the data will be written into the sheet.
            designer.Process();

            // 6. Apply conditional formatting:
            //    Highlight Salary values greater than 6000 with a yellow background.
            int cfIndex = sheet.ConditionalFormattings.Add();
            var cf = sheet.ConditionalFormattings[cfIndex];

            // Define the range that contains the Salary column (C2:C4).
            cf.AddArea(new CellArea
            {
                StartRow = 1,   // row 2 (zero‑based)
                EndRow = 3,     // row 4
                StartColumn = 2, // column C
                EndColumn = 2
            });

            // Add a CellValue condition: value > 6000.
            // The overload requires two formula strings; the second is unused for this operator.
            int conditionIndex = cf.AddCondition(
                FormatConditionType.CellValue,
                OperatorType.GreaterThan,
                "6000",
                "");

            var condition = cf[conditionIndex];
            condition.Style.BackgroundColor = Color.Yellow;

            // 7. Save the result.
            workbook.Save("AnonymousDataSmartMarkers.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
