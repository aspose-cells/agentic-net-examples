using System;
using System.Data;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains a smart marker.
            // The template should have a cell with a smart marker like "&=MyData.Formula".
            Workbook workbook = new Workbook("template.xlsx");

            // Initialize the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);

            // Prepare a data source that provides the formula string.
            DataTable dt = new DataTable("MyData");
            dt.Columns.Add("Formula", typeof(string));

            // Insert the desired Excel formula into the data source.
            // This formula will be written into the cell that contains the smart marker.
            dt.Rows.Add("=SUM(A1:A5)");

            // Assign the data source to the designer.
            designer.SetDataSource(dt);

            // Process the smart markers – the formula will be placed into the target cell.
            designer.Process();

            // Save the resulting workbook.
            workbook.Save("output.xlsx");
        }
    }
}