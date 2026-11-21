// Title: How to set a locale identifier for smart markers using WorkbookDesigner.SetVariable in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set locale for smart markers with SetVariable method | Pass culture ID to smart markers to format dates and numbers in Excel using Aspose.Cells | WorkbookDesigner SetVariable locale identifier example for date formatting in smart markers | How to control number formatting in Aspose.Cells smart markers by setting culture ID | Using SetVariable to apply US English locale to smart markers in a .NET workbook
// Tags: WorkbookDesigner SetVariable locale | smart markers culture formatting | Aspose.Cells locale identifier | C# Excel smart marker date format | C# Excel smart marker number format

using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates using WorkbookDesigner.SetVariable to pass a locale identifier (e.g., 1033 for en-US) that influences date and number formatting of smart markers in an Aspose.Cells workbook, then processing the markers and saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new empty workbook
                Workbook workbook = new Workbook();

                // Set the workbook culture (e.g., 1033 = English (United States))
                workbook.Settings.CultureInfo = new CultureInfo(1033);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Place sample smart markers
                sheet.Cells["A1"].PutValue("&=Date:DateValue");      // date smart marker
                sheet.Cells["A2"].PutValue("&=Number:NumberValue"); // number smart marker

                // Initialize the WorkbookDesigner (smart marker processor)
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // (Optional) Provide data source for the smart markers here, e.g., a DataTable or a collection
                // designer.SetDataSource("Data", dataTable);

                // Process the smart markers
                designer.Process();

                // Save the resulting workbook
                string outputPath = "SmartMarkerWithLocale.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
