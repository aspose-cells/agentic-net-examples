// Title: How to replace an array index placeholder (${i}) in a smart marker using a custom ISmartMarkerCallBack with Aspose.Cells for .NET
// AI Prompts: Write a C# ISmartMarkerCallBack that substitutes the ${i} token with the current row index during smart marker processing. | Demonstrate how to attach a custom callback to WorkbookDesigner, insert a smart marker like '&=Employees[${i}].Name' into a cell, and generate the final Excel file. | Provide sample code that creates an ArrayList data source, places a smart marker containing an array index placeholder, and saves the workbook after the callback runs.
// Common Searches: Aspose.Cells replace ${i} token in smart marker with row index | C# custom ISmartMarkerCallBack for array index placeholders | how to use WorkbookDesigner callback to modify smart marker values in .NET | smart marker with dynamic array index example Aspose.Cells
// Tags: smart marker placeholder substitution Aspose.Cells | custom ISmartMarkerCallBack C# | WorkbookDesigner callback array index | Aspose.Cells dynamic smart marker index | C# replace ${i} in smart marker

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Implement the callback that will be invoked for each smart marker occurrence
    // The example creates a workbook, inserts a smart marker containing the ${i} array index placeholder into cell A1, defines an ArrayList as the data source, implements a custom ISmartMarkerCallBack that replaces ${i} with the current row index during processing, attaches the callback to WorkbookDesigner, processes the markers, and saves the result as SmartMarkerWithIndexPlaceholder.xlsx.
    public class IndexPlaceholderCallback : ISmartMarkerCallBack
    {
        // This method is called by the smart marker processor for each marker
        public void Process(int sheetIndex, int rowIndex, int colIndex, string tableName, string columnName)
        {
            // Retrieve the worksheet and the cell that contains the smart marker
            Worksheet sheet = WorkbookDesignerInstance.Workbook.Worksheets[sheetIndex];
            Cell cell = sheet.Cells[rowIndex, colIndex];

            // Replace the placeholder ${i} with the actual row index (or any custom logic)
            // The original marker string is something like "&=Employees[${i}].Name"
            // We'll replace ${i} with the current row index (starting from 0)
            string originalValue = cell.StringValue;
            string replacedValue = originalValue.Replace("${i}", rowIndex.ToString());

            // Optionally, you can put the final value into the cell or keep the marker for further processing
            cell.PutValue(replacedValue);
        }

        // Helper to access the designer instance from the callback
        public static WorkbookDesigner WorkbookDesignerInstance { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // 2. Place a smart marker with an array index placeholder in cell A1
            // The placeholder ${i} will be replaced by the callback with the actual row index
            worksheet.Cells["A1"].PutValue("&=Employees[${i}].Name");

            // 3. Prepare a dummy data source (required for the designer to run)
            // Here we use an ArrayList with two sample employee names
            ArrayList employees = new ArrayList
            {
                new { Name = "John Doe" },
                new { Name = "Jane Smith" }
            };

            // 4. Set up the WorkbookDesigner
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook,
                LineByLine = true // Process line by line (default)
            };
            designer.SetDataSource("Employees", employees);

            // 5. Attach the callback that will replace the placeholder
            IndexPlaceholderCallback callback = new IndexPlaceholderCallback();
            // Pass the designer instance to the callback so it can access the workbook
            IndexPlaceholderCallback.WorkbookDesignerInstance = designer;
            designer.CallBack = callback;

            // 6. Process the smart markers (the callback will be invoked for each occurrence)
            designer.Process();

            // 7. Save the resulting workbook
            workbook.Save("SmartMarkerWithIndexPlaceholder.xlsx");
        }
    }
}
