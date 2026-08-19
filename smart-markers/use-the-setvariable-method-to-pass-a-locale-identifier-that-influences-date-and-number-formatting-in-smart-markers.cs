// Title: Set locale for Aspose.Cells smart markers using WorkbookDesigner.SetVariable (C#)
// Description: Demonstrates how to pass a locale identifier (LCID) to WorkbookDesigner via SetVariable (or CultureInfo) so that smart markers format dates and numbers according to the specified culture. The example creates a workbook, inserts smart markers for a date and a number, supplies a DataTable, sets the French (France) locale (LCID 1036), processes the markers, and saves the result.
// Keywords: Aspose.Cells SetVariable locale | smart markers cultureinfo | LCID formatting C# | date number localization Aspose.Cells | French locale smart markers | WorkbookDesigner SetVariable example | Excel report localization .NET
// Common Searches: Aspose.Cells set locale for smart markers | WorkbookDesigner SetVariable LCID example | How to format dates in smart markers by culture | Apply French culture to Aspose.Cells smart markers | C# smart marker localization tutorial
// Developer Intent: Pass a locale identifier to control date and number formatting of smart markers during processing.
// Use Cases: Generate Excel reports where dates appear in French format (dd/MM/yyyy) using smart markers. | Produce locale‑specific numeric values with French decimal separators (comma) in automated spreadsheets. | Create multi‑regional Excel files by setting LCID before processing smart markers with WorkbookDesigner.
// AI Prompts: Show C# code that uses WorkbookDesigner.SetVariable("LCID", 1036) to apply French formatting to smart markers. | Explain the difference between setting Workbook.Settings.CultureInfo and using SetVariable for locale handling in Aspose.Cells. | Provide examples of LCID values for different cultures and how they affect smart marker output in .NET.

using System;
using System.Data;
using System.Globalization;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerLocaleDemo
{
    // Demonstrates how to pass a locale identifier (LCID) to WorkbookDesigner via SetVariable (or CultureInfo) so that smart markers format dates and numbers according to the specified culture. The example creates a workbook, inserts smart markers for a date and a number, supplies a DataTable, sets the French (France) locale (LCID 1036), processes the markers, and saves the result.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Insert smart markers into cells.
                // &=DateField; will be replaced by the date value.
                // &=NumberField; will be replaced by the numeric value.
                sheet.Cells["A1"].PutValue("&=DateField;");
                sheet.Cells["B1"].PutValue("&=NumberField;");

                // Prepare a data source with a date and a number.
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("DateField", typeof(DateTime));
                dt.Columns.Add("NumberField", typeof(double));
                dt.Rows.Add(new DateTime(2023, 12, 31), 12345.67);

                // Set the desired locale (LCID 1036 = French (France)).
                workbook.Settings.CultureInfo = new CultureInfo(1036);

                // Create a WorkbookDesigner to process smart markers.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // Set the data source for the smart markers.
                designer.SetDataSource(dt);

                // Process the smart markers with the provided data source and locale.
                designer.Process();

                // Define output file path.
                string outputPath = "SmartMarkerLocaleResult.xlsx";

                // Ensure the output directory exists before saving.
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the resulting workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
