using System;
using Aspose.Cells;

namespace SmartMarkerArrayIndexDemo
{
    class Program
    {
        static void Main()
        {
            // Load the Excel template that contains smart markers.
            // The template should have a cell with a smart marker like: &=Names[1]
            // This marker will be replaced by the element at index 1 of the array (second element).
            Workbook workbook = new Workbook("Template.xlsx");

            // Initialize the WorkbookDesigner and associate it with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = workbook
            };

            // Prepare a simple array data source.
            // The smart marker will access elements by their zero‑based index.
            string[] names = new string[] { "Alice", "Bob", "Charlie", "Diana" };

            // Register the array as a data source with the name "Names".
            designer.SetDataSource("Names", names);

            // Process the smart markers in the workbook.
            // This will replace the marker &=Names[1] with "Bob".
            designer.Process();

            // Save the populated workbook.
            workbook.Save("Result.xlsx");
        }
    }
}