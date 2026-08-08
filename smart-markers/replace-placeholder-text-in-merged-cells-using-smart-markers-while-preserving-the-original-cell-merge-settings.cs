// Title: C# – Replace Smart Marker in Merged Cells with Aspose.Cells while Keeping Merge Layout
// Description: Demonstrates how to merge a range (A1:B2), insert a smart marker "&=Data.Name", bind a List<Person> as the data source, process the marker with WorkbookDesigner, and save the workbook, all while preserving the original merged‑cell formatting.
// Keywords: Aspose.Cells smart markers C# | replace placeholder merged cells | preserve merge settings Aspose | WorkbookDesigner merge range | dynamic header with smart marker
// Common Searches: Aspose.Cells keep merged cells after smart marker processing | C# replace smart marker inside merged range | how to preserve cell merge when using smart markers | smart marker merge example Aspose.Cells .NET
// Developer Intent: Insert a smart marker into a merged cell, bind data, process the marker, and retain the original merged‑cell configuration.
// Use Cases: Generate personalized reports where a merged header cell displays a name or title via a smart marker. | Create invoice or certificate templates with merged title rows that automatically fill with dynamic data without breaking layout.
// AI Prompts: Write C# code that merges A1:B2, adds the smart marker '&=Employee.FullName', binds a List<Employee>, processes the markers with WorkbookDesigner, and saves the file while preserving the merge. | Explain how WorkbookDesigner maintains merged cell ranges when processing smart markers and list any properties that affect this behavior. | Provide a step‑by‑step tutorial for replacing placeholder text in merged cells using Aspose.Cells smart markers, including loading an existing workbook and handling multiple data rows.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerMergeDemo
{
    // Simple data class for the smart marker
    // Demonstrates how to merge a range (A1:B2), insert a smart marker "&=Data.Name", bind a List<Person> as the data source, process the marker with WorkbookDesigner, and save the workbook, all while preserving the original merged‑cell formatting.
    public class Person
    {
        public string Name { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load or create a workbook (here we create a new one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Merge a range of cells (A1:B2) – this will be the placeholder area
            cells.Merge(0, 0, 2, 2); // rows are zero‑based, totalRows/totalColumns are 1‑based

            // Put a smart marker inside the merged cell.
            // The marker syntax "&=Data.Name" tells the designer to replace it with the Name property of the data source named "Data".
            cells["A1"].PutValue("&=Data.Name");

            // Prepare the data source – a list with a single Person object
            List<Person> data = new List<Person>
            {
                new Person { Name = "John Doe" }
            };

            // Set up the WorkbookDesigner, assign the workbook and the data source
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };
            designer.SetDataSource("Data", data);

            // Process the smart markers. The merge settings are preserved automatically.
            designer.Process();

            // Save the result
            workbook.Save("MergedSmartMarkerResult.xlsx");
        }
    }
}
