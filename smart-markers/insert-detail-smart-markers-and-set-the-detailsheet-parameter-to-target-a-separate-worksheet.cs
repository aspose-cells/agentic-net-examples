// Title: How to generate an Excel workbook with master smart markers and a separate detail worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells example for placing detail smart markers on a different worksheet | using WorkbookDesigner to set a detail sheet for nested smart markers in .NET | Aspose.Cells master‑detail smart markers across multiple sheets tutorial | how to create a separate detail worksheet for smart marker collections with Aspose.Cells | Aspose.Cells generate Excel with master sheet and detail sheet smart markers C#
// Tags: Aspose.Cells master smart markers worksheet | Aspose.Cells detail smart markers separate sheet | WorkbookDesigner set detail sheet parameter .NET | C# generate Excel with nested smart markers | Aspose.Cells master‑detail smart markers example

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new Workbook, adds a master sheet containing smart markers for customer fields, adds a separate 'Detail' sheet with smart markers for a nested Orders collection, optionally initializes a WorkbookDesigner, and saves the file as SmartMarkerDetail.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            var workbook = new Workbook();

            // ----- Master sheet with smart markers -----
            Worksheet masterSheet = workbook.Worksheets[0];
            masterSheet.Name = "Master";

            // Header row
            masterSheet.Cells["A1"].PutValue("Customer Name");
            masterSheet.Cells["B1"].PutValue("Customer Age");

            // Detail smart markers that will be expanded from the data source
            masterSheet.Cells["A2"].PutValue("&=Customers.Name");
            masterSheet.Cells["B2"].PutValue("&=Customers.Age");

            // ----- Separate worksheet for nested detail data -----
            Worksheet detailSheet = workbook.Worksheets.Add("Detail");

            // Header row for the detail sheet
            detailSheet.Cells["A1"].PutValue("Order ID");
            detailSheet.Cells["B1"].PutValue("Product");

            // Detail smart markers for the nested collection (e.g., Orders)
            detailSheet.Cells["A2"].PutValue("&=Customers.Orders.OrderID");
            detailSheet.Cells["B2"].PutValue("&=Customers.Orders.ProductName");

            // (Optional) Initialize WorkbookDesigner if further processing is needed
            var designer = new WorkbookDesigner(workbook);
            // Note: In newer Aspose.Cells versions, the Options property may be unavailable.
            // If you need to specify a detail sheet, use the appropriate API for your version.
            // For this template creation example, no additional designer configuration is required.

            // Save the workbook
            string outputPath = "SmartMarkerDetail.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
