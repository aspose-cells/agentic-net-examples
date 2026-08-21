// Title: Apply a custom style to smart‑marker output cells using Aspose.Cells for .NET
// Description: Loads a workbook template with smart markers, binds a DataTable as the data source, processes the markers with WorkbookDesigner, creates a Calibri 11 pt style with a light‑yellow background, iterates over all used cells, applies the style only to cells that contain data after processing while preserving any existing formatting, and saves the styled workbook.
// Keywords: Aspose.Cells | C# | .NET | smart markers | custom cell style | WorkbookDesigner | SetStyle | StyleFlag | iterate cells | preserve formatting | Excel template | data source | report generation
// Common Searches: how to style smart marker results in Aspose.Cells | apply formatting only to populated cells after WorkbookDesigner process | preserve existing cell formatting when adding a style with Aspose.Cells | C# example for custom style on smart marker output | Aspose.Cells set style after smart marker expansion
// Developer Intent: Add a predefined formatting style to every cell that receives data from smart markers after the workbook has been processed.
// Use Cases: Generate a uniform look for rows created by smart markers in an employee report. | Highlight all non‑empty cells after smart‑marker expansion while keeping any manual formatting intact. | Automate styling of a templated Excel file after data insertion for consistent branding.
// AI Prompts: Write C# code that creates an Aspose.Cells Style and applies it to all cells with values after WorkbookDesigner processes smart markers. | Show how to modify the cell‑iteration loop to exclude header rows while still styling smart‑marker‑filled cells. | Provide a snippet that uses StyleFlag to change only font color and background for cells populated by smart markers.

using System;
using System.Data;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace SmartMarkerStyleDemo
{
    // Loads a workbook template with smart markers, binds a DataTable as the data source, processes the markers with WorkbookDesigner, creates a Calibri 11 pt style with a light‑yellow background, iterates over all used cells, applies the style only to cells that contain data after processing while preserving any existing formatting, and saves the styled workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the template workbook that contains smart markers
                string templatePath = "TemplateWithSmartMarkers.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file not found: {templatePath}");
                    return;
                }

                // Load the template workbook
                Workbook workbook = new Workbook(templatePath);

                // Initialize WorkbookDesigner with the loaded workbook
                WorkbookDesigner designer = new WorkbookDesigner(workbook);

                // ----- Prepare a simple data source -----
                DataTable dt = new DataTable("Employees");
                dt.Columns.Add("Name", typeof(string));
                dt.Columns.Add("Age", typeof(int));
                dt.Columns.Add("Department", typeof(string));

                dt.Rows.Add("John Doe", 30, "Sales");
                dt.Rows.Add("Jane Smith", 28, "Marketing");

                // Set the data source for the smart markers
                designer.SetDataSource(dt);

                // Process all smart markers in the workbook
                designer.Process();

                // ----- Create a custom style to be applied to smart‑marker cells -----
                Style customStyle = workbook.CreateStyle();
                customStyle.Font.Name = "Calibri";
                customStyle.Font.Size = 11;
                customStyle.Font.Color = Color.Black;
                customStyle.ForegroundColor = Color.LightYellow;
                customStyle.Pattern = BackgroundType.Solid;

                // Use a StyleFlag to apply all formatting properties
                StyleFlag flag = new StyleFlag { All = true };

                // Apply the custom style only to cells that were populated by smart markers.
                // After processing, these cells contain data (i.e., they are not empty).
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Determine if the cell contains any value after smart‑marker processing
                        bool hasValue = cell.Value != null && cell.Type != CellValueType.IsNull;

                        if (hasValue)
                        {
                            // Apply the custom style while preserving any explicitly set formatting
                            cell.SetStyle(customStyle, true);
                        }
                    }
                }

                // Save the workbook with the applied formatting
                string outputPath = "SmartMarkerStyledOutput.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
