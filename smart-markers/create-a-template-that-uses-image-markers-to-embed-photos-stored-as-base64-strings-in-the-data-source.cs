// Title: How to embed a Base64‑encoded picture into an Excel file using Aspose.Cells smart image markers in C#
// AI Prompts: Write C# code that creates an Excel workbook, places a smart image marker ${Photo} in a cell, decodes a Base64 string to a byte array, and binds it to a DataTable for WorkbookDesigner processing. | Show how to use Aspose.Cells WorkbookDesigner to set a DataTable with a byte[] column as the data source and replace the image marker with the decoded picture. | Demonstrate converting a Base64 image string to a byte[] and inserting it into an Excel template via a smart image marker using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# insert image from Base64 string using smart markers | How to bind a byte array column to WorkbookDesigner for picture insertion | Excel template with ${Photo} marker and Base64 image data in C# | Convert Base64 to byte[] and display as picture in Aspose.Cells workbook | Smart image marker example with DataTable in Aspose.Cells .NET
// Tags: Aspose.Cells smart image marker byte array | C# embed base64 picture Excel | WorkbookDesigner bind image data source | Excel template image marker Aspose.Cells | convert base64 to image Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;

// The program creates a new workbook, adds a smart image marker ${Photo} in cell A1, decodes a Base64 string to a byte array, stores it in a DataTable column of type byte[], sets the DataTable as the data source for WorkbookDesigner, processes the template to replace the marker with the actual picture, and saves the result as ImageMarkerTemplate.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Insert an image marker in cell A1. The marker name (Photo) will be used to bind the image data.
        sheet.Cells["A1"].PutValue("${Photo}");

        // ----- Prepare data source -----
        // Example base64 string of an image (replace with actual data)
        string base64Image = "iVBORw0KGgoAAAANSUhEUgAAAAUA" +
                             "AAAFCAYAAACNbyblAAAAHElEQVQI12P4" +
                             "//8/w38GIAXDIBKE0DHxgljNBAAO" +
                             "9TXL0Y4OHwAAAABJRU5ErkJggg==";

        // Convert the base64 string to a byte array
        byte[] imageBytes = Convert.FromBase64String(base64Image);

        // Create a DataTable with a column of type byte[] to hold the image data
        DataTable data = new DataTable();
        data.Columns.Add("Photo", typeof(byte[]));
        DataRow dr = data.NewRow();
        dr["Photo"] = imageBytes;
        data.Rows.Add(dr);

        // ----- Bind data source and process the template -----
        WorkbookDesigner designer = new WorkbookDesigner(workbook);
        designer.SetDataSource(data);
        designer.Process(); // This replaces the ${Photo} marker with the actual image

        // Save the resulting workbook
        workbook.Save("ImageMarkerTemplate.xlsx");
    }
}
