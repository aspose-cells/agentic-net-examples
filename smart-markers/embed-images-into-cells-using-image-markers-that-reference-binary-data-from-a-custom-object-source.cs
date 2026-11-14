// Title: Embedding a logo from a byte array into an Excel cell with Aspose.Cells smart image markers (C#)
// Common Searches: C# Aspose.Cells smart marker to insert image from byte[] into a specific cell | Using WorkbookDesigner to replace image markers with binary data from a DataSet | Load PNG file as byte array and bind it to a smart image marker in Aspose.Cells | How to handle missing image file when using smart image markers in Aspose.Cells | Example of embedding a logo from a DataSet into cell B2 with Aspose.Cells
// Tags: Aspose.Cells smart image marker | WorkbookDesigner bind DataSet image | embed byte array image Excel C# | replace image marker binary data | load png as byte[] for Aspose.Cells

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample creates a new workbook, places the smart image marker ${ImageDataSource.Logo} in cell B2, reads logo.png into a byte array, constructs a DataSet with a table named ImageDataSource containing the byte array, and uses WorkbookDesigner to bind the DataSet and process the marker. The resulting workbook is saved as Output.xlsx, with console warnings for missing files or save errors.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert an image marker in cell B2.
                // For DataSet data source the marker must include the table name.
                sheet.Cells["B2"].PutValue("${ImageDataSource.Logo}");

                // Prepare image data
                const string logoPath = "logo.png";
                byte[] logoBytes = null;
                if (File.Exists(logoPath))
                {
                    try
                    {
                        logoBytes = File.ReadAllBytes(logoPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Failed to read image file '{logoPath}'. {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Image file '{logoPath}' not found. Marker will be left empty.");
                }

                // Build a DataSet with a single table containing the image bytes
                DataSet ds = new DataSet();
                DataTable dt = new DataTable("ImageDataSource");
                dt.Columns.Add("Logo", typeof(byte[]));
                dt.Rows.Add(logoBytes ?? (object)DBNull.Value);
                ds.Tables.Add(dt);

                // Use WorkbookDesigner to replace the marker with the actual image
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                designer.SetDataSource(ds);
                designer.Process();

                // Save the resulting workbook
                const string outputPath = "Output.xlsx";
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
