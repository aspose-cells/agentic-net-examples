// Title: How to embed an image from a file path into an Excel cell using Aspose.Cells smart markers in C#
// AI Prompts: Write C# code that places a smart marker "&=Image" in a worksheet cell and uses WorkbookDesigner to replace it with an image loaded from a file path supplied in a DataTable. | Show how to bind a DataTable column that holds image file paths to a smart marker and generate an Excel workbook with those pictures inserted. | Modify the example to process multiple rows, inserting a different picture for each row using the same smart marker definition.
// Common Searches: aspocells smart marker insert image from file path c# example | c# workbookdesigner replace smart marker with picture using datatable | how to bind image column to smart marker in aspocells | dynamic image insertion into excel using aspocells smart markers
// Tags: Aspose.Cells WorkbookDesigner image smart marker | C# embed picture into Excel via smart marker | smart marker image file path binding | dynamic picture insertion Aspose.Cells | Excel generation with image column DataTable

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsSmartMarkerImageExample
{
    // The program creates a workbook, adds a smart marker "&=Image" to cell A1, supplies a DataTable with an "Image" column containing a file path, processes the marker with WorkbookDesigner, and saves the result as SmartMarkerImageOutput.xlsx, embedding the referenced picture into the worksheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a smart marker that expects an image.
            // The marker name "Image" will be matched with the column name in the data source.
            sheet.Cells["A1"].PutValue("&=Image");

            // Prepare a data source with a column named "Image" containing the file path.
            DataTable dt = new DataTable();
            dt.Columns.Add("Image", typeof(string));

            // Add a row with the image file path (adjust the path as needed).
            DataRow row = dt.NewRow();
            row["Image"] = "C:\\Images\\sample.jpg"; // Path to the image file
            dt.Rows.Add(row);

            // Use WorkbookDesigner to process the smart marker.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(dt);
            designer.Process();

            // Save the resulting workbook.
            workbook.Save("SmartMarkerImageOutput.xlsx");
        }
    }
}
