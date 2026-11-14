// Title: Enable DetailLink and add row-level hyperlinks from a master worksheet to a detail worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create hyperlink from master row to detail worksheet | Enable DetailLink smart marker to link master and detail sheets in .NET | Add row-level worksheet hyperlinks using Aspose.Cells API | Programmatically set ScreenTip for Excel hyperlinks with Aspose.Cells | Generate master-detail Excel file with clickable links in C#
// Tags: Aspose.Cells DetailLink API | C# master-detail worksheet hyperlink | Excel hyperlink creation Aspose.Cells | ScreenTip property Aspose.Cells Hyperlink | Generate linked master-detail workbook C#

using System;
using Aspose.Cells;

namespace AsposeCellsDetailLinkExample
{
    // The example creates a workbook with a master and a detail worksheet, populates sample data, then loops through master rows to insert hyperlinks in column A that point to the corresponding rows on the detail sheet. Each hyperlink receives a ScreenTip, and the workbook is saved as an .xlsx file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook wb = new Workbook();

                // Add a master worksheet
                int masterIndex = wb.Worksheets.Add();
                Worksheet master = wb.Worksheets[masterIndex];
                master.Name = "Master";

                // Add a detail worksheet
                int detailIndex = wb.Worksheets.Add();
                Worksheet detail = wb.Worksheets[detailIndex];
                detail.Name = "Detail";

                // Populate master sheet (example data)
                master.Cells["A1"].PutValue("ID");
                master.Cells["B1"].PutValue("Name");
                master.Cells["A2"].PutValue(1);
                master.Cells["B2"].PutValue("Item 1");
                master.Cells["A3"].PutValue(2);
                master.Cells["B3"].PutValue("Item 2");

                // Populate detail sheet (example data)
                detail.Cells["A1"].PutValue("MasterID");
                detail.Cells["B1"].PutValue("DetailInfo");
                detail.Cells["A2"].PutValue(1);
                detail.Cells["B2"].PutValue("Detail for Item 1");
                detail.Cells["A3"].PutValue(2);
                detail.Cells["B3"].PutValue("Detail for Item 2");

                // Enable DetailLink by adding hyperlinks from each master row to its corresponding detail row
                // Row indices are zero‑based; column A (index 0) will hold the hyperlink
                for (int i = 1; i <= 2; i++) // rows 2 and 3 in the sheet
                {
                    // Build the address that points to the matching row in the Detail sheet
                    // Example: 'Detail'!A2, 'Detail'!A3, etc.
                    string address = $"'{detail.Name}'!A{i + 1}";

                    // Add the hyperlink (returns the index of the newly added hyperlink)
                    int hlIndex = master.Hyperlinks.Add(i, 0, 1, 1, address);

                    // Retrieve the Hyperlink object to set additional properties
                    Hyperlink hl = master.Hyperlinks[hlIndex];
                    hl.ScreenTip = $"Go to detail for ID {master.Cells[i, 0].IntValue}";
                }

                // Save the workbook
                string outputPath = "DetailLinkEnabled.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
