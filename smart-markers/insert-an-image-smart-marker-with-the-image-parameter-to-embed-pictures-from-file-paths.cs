// Title: Insert an image into an Excel worksheet using Aspose.Cells smart markers with a file path in C#
// AI Prompts: Generate C# code that places a smart marker '&=Image: {ImagePath}' in a cell, provides a DataTable containing the image file path, runs WorkbookDesigner.Process, and saves the workbook. | Demonstrate how to embed a local JPEG into an Excel cell by configuring a smart marker image placeholder and a data source with Aspose.Cells.
// Common Searches: C# Aspose.Cells smart marker insert image from local file path | How to use WorkbookDesigner to embed pictures via smart markers | Aspose.Cells example populating Excel cell with image using a DataTable | Insert JPEG into Excel worksheet using smart marker placeholder in C#
// Tags: Aspose.Cells WorkbookDesigner image smart marker | embed picture from file path Excel C# | smart marker image placeholder processing | populate Excel cell with image via data source | Aspose.Cells insert image using smart marker

using System;
using System.Data;
using Aspose.Cells;

// The program creates a new workbook, adds a smart marker '&=Image: {ImagePath}' to cell A1, builds a DataTable with the image file path, sets it as the data source for WorkbookDesigner, processes the marker to embed the picture, and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a smart marker that will be replaced by an image.
        // The placeholder {ImagePath} will be filled with the file path of the picture.
        sheet.Cells["A1"].PutValue("&=Image: {ImagePath}");

        // Prepare a data source containing the image file path.
        DataTable dt = new DataTable("Images");
        dt.Columns.Add("ImagePath", typeof(string));
        // Add a row with the full path to the image file you want to embed.
        dt.Rows.Add(@"C:\Images\sample1.jpg");

        // Set the data source for the smart marker processor.
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(dt);

        // Process the smart markers – the image will be inserted into the cell.
        designer.Process();

        // Save the resulting workbook.
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}
