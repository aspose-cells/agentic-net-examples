using System;
using System.Collections;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the template workbook that contains smart markers
        Workbook workbook = new Workbook("Template.xlsx");

        // Initialize the WorkbookDesigner with the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner(workbook);

        // Prepare custom data as an ArrayList of object arrays (rows)
        ArrayList data = new ArrayList();
        data.Add(new object[] { "Name", "Age", "City" });          // Header row (optional)
        data.Add(new object[] { "John", 30, "New York" });
        data.Add(new object[] { "Anna", 25, "London" });
        data.Add(new object[] { "Mike", 35, "Sydney" });

        // Create an ICellsDataTable instance from the custom data
        ICellsDataTable customSource = workbook.CellsDataTableFactory.GetInstance(data, true);

        // Bind the custom data source to the designer.
        // The name "People" must match the smart marker name used in the template (e.g., &People.Name)
        designer.SetDataSource("People", customSource);

        // Process the smart markers and populate the worksheet with data
        designer.Process();

        // Save the populated workbook
        workbook.Save("Result.xlsx");
    }
}