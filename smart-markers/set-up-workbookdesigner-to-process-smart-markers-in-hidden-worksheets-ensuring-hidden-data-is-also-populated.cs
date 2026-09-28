// Title: How to use Aspose.Cells WorkbookDesigner to populate smart markers on a hidden worksheet in C#
// AI Prompts: Instantiate WorkbookDesigner, load a workbook, hide the target worksheet, add smart markers, bind a List<Employee> to the 'Employees' prefix, set UpdateReference = true, and call Process() to fill the hidden sheet. | Create a hidden worksheet named "HiddenData", place smart markers such as "&=Employees.Name" and "&=Employees.Age", bind the employee collection, enable reference updates, and save the processed file. | Use Aspose.Cells to ensure hidden worksheets are included in smart‑marker processing by configuring WorkbookDesigner, binding object data, and generating the final Excel output.
// Common Searches: Aspose.Cells C# populate smart markers on a hidden worksheet | WorkbookDesigner process hidden sheet smart markers .NET | How to bind a list of objects to smart markers in a hidden Excel sheet using Aspose.Cells | Enable UpdateReference to refresh formulas after processing hidden smart markers in C# | Create hidden worksheet with smart markers and fill data using Aspose.Cells Designer
// Tags: WorkbookDesigner hidden sheet data binding | process smart markers concealed worksheet | Aspose.Cells update reference after smart marker fill | C# populate hidden Excel sheet with object list | smart marker prefix Employees binding

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads a template workbook, ensures a hidden worksheet named "HiddenData" exists, places smart markers, binds a List<Employee> to the "Employees" prefix with WorkbookDesigner, sets UpdateReference to true, processes all smart markers including those on hidden sheets, and saves the result to output.xlsx.
class ProcessHiddenSmartMarkers
{
    static void Main()
    {
        // Load the template workbook that contains smart markers in a hidden worksheet
        Workbook workbook = new Workbook("template.xlsx");

        // Ensure the worksheet that holds the hidden data exists and is hidden
        Worksheet hiddenSheet = workbook.Worksheets["HiddenData"];
        if (hiddenSheet == null)
        {
            hiddenSheet = workbook.Worksheets.Add("HiddenData");
        }
        hiddenSheet.IsVisible = false; // hide the worksheet

        // Example smart markers placed in the hidden sheet (for illustration)
        // hiddenSheet.Cells["A1"].PutValue("&=Employees.Name");
        // hiddenSheet.Cells["B1"].PutValue("&=Employees.Age");

        // Prepare a data source that matches the smart marker prefix "Employees"
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John Doe", Age = 30 },
            new Employee { Name = "Jane Smith", Age = 28 }
        };

        // Initialize WorkbookDesigner and assign the loaded workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = workbook;

        // Bind the data source to the smart marker name
        designer.SetDataSource("Employees", employees);

        // Update references in other worksheets (including hidden ones) after processing
        designer.UpdateReference = true;

        // Process all smart markers, hidden worksheets included
        designer.Process();

        // Save the processed workbook
        workbook.Save("output.xlsx");
    }

    // Simple POCO class representing the data source
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
