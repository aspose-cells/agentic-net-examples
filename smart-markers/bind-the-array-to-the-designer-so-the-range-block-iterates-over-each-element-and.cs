using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsArrayBindingDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains a RANGE smart marker block.
            // The range should be named (e.g., "_CellsSmartMarkers") in the template.
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare the array that will be bound to the smart marker.
            // The name "MyArray" must match the variable used in the smart marker (e.g., &=$MyArray).
            string[] names = new string[] { "Alice", "Bob", "Charlie" };

            // Bind the array to the designer. This will cause the RANGE block to repeat
            // for each element of the array during processing.
            designer.SetDataSource("MyArray", names);

            // Process the smart markers, filling the range with the array data.
            designer.Process();

            // Save the populated workbook.
            workbook.Save("output.xlsx");
        }
    }
}