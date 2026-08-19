// Title: C# – Process Smart Markers in Hidden Worksheets with Aspose.Cells WorkbookDesigner
// Description: Demonstrates how to add a hidden worksheet, place smart markers, bind a List<Employee> data source, enable UpdateReference, and call WorkbookDesigner.Process() so hidden‑sheet markers are populated and saved to an Excel file.
// Keywords: Aspose.Cells | WorkbookDesigner | smart markers | hidden worksheet | C# | SetDataSource | UpdateReference | Excel report generation | populate hidden sheet | Aspose.Cells .NET example
// Common Searches: Aspose.Cells process smart markers in hidden sheet | WorkbookDesigner hidden worksheet example C# | populate hidden Excel sheet with smart markers | UpdateReference smart markers Aspose.Cells | set datasource for smart markers hidden tab
// Developer Intent: Automatically fill smart markers that reside on a hidden worksheet using WorkbookDesigner.
// Use Cases: Create a clean report layout where calculation tables are hidden but still driven by smart markers. | Store lookup or configuration data in hidden tabs and populate them at runtime without exposing them to end users. | Generate templates with internal data sections that remain invisible while still participating in smart‑marker processing.
// AI Prompts: Generate C# code that uses Aspose.Cells WorkbookDesigner to process smart markers on both visible and hidden worksheets, ensuring hidden data is filled. | Explain how the UpdateReference property influences formulas that reference hidden worksheets after smart‑marker processing. | Show how to bind a List<T> collection to a smart‑marker name and process it on a hidden sheet with Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// Demonstrates how to add a hidden worksheet, place smart markers, bind a List<Employee> data source, enable UpdateReference, and call WorkbookDesigner.Process() so hidden‑sheet markers are populated and saved to an Excel file.
class ProcessHiddenSmartMarkers
{
    static void Main()
    {
        // Create a new workbook that will serve as the template
        Workbook wb = new Workbook();

        // Add a hidden worksheet and place smart markers in it
        Worksheet hiddenSheet = wb.Worksheets.Add("HiddenData");
        hiddenSheet.IsVisible = false; // Hide the worksheet
        hiddenSheet.Cells["A1"].PutValue("&=Employees.Name"); // Smart marker for Name
        hiddenSheet.Cells["B1"].PutValue("&=Employees.Age");  // Smart marker for Age

        // Add a visible worksheet for reference (optional)
        Worksheet visibleSheet = wb.Worksheets[0];
        visibleSheet.Name = "Report";
        visibleSheet.Cells["A1"].PutValue("Employee Report");
        visibleSheet.Cells["A2"].PutValue("Data is populated in the hidden sheet.");

        // Prepare a data source that matches the smart marker name
        List<Employee> employees = new List<Employee>
        {
            new Employee { Name = "John Doe", Age = 30 },
            new Employee { Name = "Jane Smith", Age = 28 }
        };

        // Initialize WorkbookDesigner and assign the workbook
        WorkbookDesigner designer = new WorkbookDesigner();
        designer.Workbook = wb;

        // Optional: update references in other sheets after processing
        designer.UpdateReference = true;

        // Bind the data source to the name used in the smart markers
        designer.SetDataSource("Employees", employees);

        // Process all smart markers, including those in hidden worksheets
        designer.Process();

        // Save the processed workbook
        wb.Save("ProcessedHiddenSmartMarkers.xlsx");
    }

    // Simple POCO class representing employee data
    public class Employee
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
