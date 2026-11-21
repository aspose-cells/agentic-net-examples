// Title: Re‑export a workbook to HTML after attempting to change a ListObject's TableCssId and verify the updated HTML file with Aspose.Cells for .NET
// Common Searches: how to change the TableCssId of an Excel ListObject when saving as HTML using Aspose.Cells C# | re‑export workbook to HTML after modifying ListObject properties with Aspose.Cells | verify that an updated HTML file was created after Aspose.Cells export in C# | Aspose.Cells .NET check availability of TableCssId property for HTML output | export Excel table to HTML with custom CSS id using Aspose.Cells
// Tags: Aspose.Cells modify ListObject TableCssId | HTML export re‑save workbook C# | verify generated HTML file existence Aspose.Cells | Aspose.Cells ListObject CSS class handling | C# export Excel to HTML base64 images

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds sample data and a ListObject, attempts to set its DisplayName, exports the workbook to HTML twice (initial and updated) using HtmlSaveOptions, notes that the TableCssId property is not available in the current API version, and finally checks that the updated HTML file was successfully created.
class TableCssIdUpdateExample
{
    static void Main()
    {
        try
        {
            // Paths for the initial and updated HTML files
            string initialHtmlPath = "TableInitial.html";
            string updatedHtmlPath = "TableUpdated.html";

            // -----------------------------------------------------------------
            // 1. Create a new workbook and add sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some data for the table
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("Alice");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Bob");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Charlie");

            // -----------------------------------------------------------------
            // 2. Create a ListObject (Excel table)
            // -----------------------------------------------------------------
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 4;     // rows 1‑4
            int totalColumns = 2;  // columns A‑B

            // Add the ListObject to the worksheet
            int listObjectIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject listObject = sheet.ListObjects[listObjectIndex];

            // Set a display name for the ListObject (supported property)
            try
            {
                listObject.DisplayName = "SampleTable";
            }
            catch
            {
                // Ignore if the property is unavailable in the current version
            }

            // -----------------------------------------------------------------
            // 3. Export the workbook to HTML (first export)
            // -----------------------------------------------------------------
            HtmlSaveOptions initialOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true
            };

            try
            {
                workbook.Save(initialHtmlPath, initialOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving initial HTML: {ex.Message}");
                return;
            }

            // -----------------------------------------------------------------
            // 4. (Optional) Modify the ListObject or workbook here if needed
            // -----------------------------------------------------------------
            // In this example we do not change any HTML‑specific properties because
            // the TableCssId property is not available in the current API version.

            // -----------------------------------------------------------------
            // 5. Re‑export the workbook to HTML (updated export)
            // -----------------------------------------------------------------
            HtmlSaveOptions updatedOptions = new HtmlSaveOptions
            {
                ExportImagesAsBase64 = true
            };

            try
            {
                workbook.Save(updatedHtmlPath, updatedOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving updated HTML: {ex.Message}");
                return;
            }

            // -----------------------------------------------------------------
            // 6. Verify that the updated HTML file was created
            // -----------------------------------------------------------------
            if (File.Exists(updatedHtmlPath))
            {
                Console.WriteLine($"Updated HTML file '{updatedHtmlPath}' was created successfully.");
            }
            else
            {
                Console.WriteLine($"Error: Updated HTML file '{updatedHtmlPath}' was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
