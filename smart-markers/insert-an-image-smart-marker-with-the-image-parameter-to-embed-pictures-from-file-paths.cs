// Title: Insert an image into an Excel worksheet using a file‑path smart marker with Aspose.Cells for .NET (C#)
// Common Searches: how to use a file path in an image smart marker with Aspose.Cells C# | Aspose.Cells embed picture from local file using smart markers | C# example for inserting images via WorkbookDesigner and smart markers | create Excel template with image smart marker syntax Aspose.Cells | process image smart markers from DataTable containing file paths
// Tags: image smart marker with file path Aspose.Cells | WorkbookDesigner embed picture from DataTable | process smart markers for images C# | Excel template image smart marker syntax | Aspose.Cells data source image column handling

using System;
using System.Data;
using System.IO;
using Aspose.Cells;

// The sample loads or creates an Excel template that contains an image smart marker, builds a DataTable with a column holding the image file path, sets this table as the data source for a WorkbookDesigner, processes the smart marker to embed the image into the worksheet, and saves the resulting workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string templatePath = "TemplateWithImageSmartMarker.xlsx";
            const string resultPath = "ResultWithEmbeddedImage.xlsx";
            const string imagePath = @"C:\Images\SampleImage.png";

            // Ensure the template exists; if not, create a minimal one with an image smart marker.
            Workbook workbook;
            if (File.Exists(templatePath))
            {
                workbook = new Workbook(templatePath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                // Insert the image smart marker syntax into cell A1.
                sheet.Cells["A1"].PutValue("&=\"ImagePath\"");
                // Optionally save the generated template for future runs.
                workbook.Save(templatePath);
            }

            // Prepare the data source with a column name matching the smart marker.
            DataTable imageTable = new DataTable();
            imageTable.Columns.Add("ImagePath", typeof(string));

            // Verify the image file exists before adding it to the table.
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Warning: Image file not found at '{imagePath}'. The smart marker will be left empty.");
                imageTable.Rows.Add(DBNull.Value);
            }
            else
            {
                DataRow row = imageTable.NewRow();
                row["ImagePath"] = imagePath;
                imageTable.Rows.Add(row);
            }

            // Apply the data source to the workbook designer and process smart markers.
            WorkbookDesigner designer = new WorkbookDesigner(workbook);
            designer.SetDataSource(imageTable);
            designer.Process();

            // Save the final workbook.
            workbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
