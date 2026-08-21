// Title: Aspose.Cells for .NET – Create an Excel table and a named range for the data body only
// Description: C# example that builds a workbook, adds a ListObject (table) with headers, extracts the table's DataRange (excluding the header), assigns it the name "EmployeeData", and saves the file as TableWithNamedDataRange.xlsx.
// Keywords: Aspose.Cells C# table named range | Excel ListObject data range Aspose | create named range without header Aspose.Cells | Aspose.Cells .NET table body range | C# Excel named range from table data
// Common Searches: Aspose.Cells create named range for table data only | C# get ListObject data range excluding header | define workbook‑level named range from Excel table Aspose | how to name the data body of a table in Aspose.Cells
// Developer Intent: Generate an Excel table and a workbook‑level named range that points only to the table’s data rows, not the header row.
// Use Cases: Use the named range as a chart source so the header isn’t plotted. | Reference the range in formulas or data validation that should apply only to data rows. | Copy or export the data‑only range to another workbook while preserving structure.
// AI Prompts: Provide C# code that adds a ListObject, retrieves its DataRange, and creates a named range that excludes the header using Aspose.Cells. | Show how to link the named range "EmployeeData" to a chart series in Aspose.Cells for .NET. | Explain how to modify or delete the named range created from a table’s data body after the workbook is saved.

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Tables;
using AsposeRange = Aspose.Cells.Range;

// C# example that builds a workbook, adds a ListObject (table) with headers, extracts the table's DataRange (excluding the header), assigns it the name "EmployeeData", and saves the file as TableWithNamedDataRange.xlsx.
class Program
{
    static void Main()
    {
        // Register code page provider (required for some locales)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Name");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue("John");
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue("Mary");
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue("Steve");

            // Add a ListObject (table) covering the range A1:B4; the first row has headers
            int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "EmployeeTable";

            // Get the data range of the table (excludes the header row)
            AsposeRange dataRange = table.DataRange;

            // Create a named range that refers only to the data body
            dataRange.Name = "EmployeeData";

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithNamedDataRange.xlsx");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
