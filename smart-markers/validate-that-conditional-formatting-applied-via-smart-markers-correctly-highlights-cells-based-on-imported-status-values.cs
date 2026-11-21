// Title: C# program to validate conditional formatting colors applied by smart markers in an Excel workbook using Aspose.Cells
// Common Searches: how to programmatically verify conditional formatting colors after processing smart markers with Aspose.Cells in C# | C# Aspose.Cells check cell background color for status values using WorkbookDesigner | validate Excel conditional formatting applied by smart markers using Aspose.Cells .NET | Aspose.Cells example to test conditional formatting results for a status column in a generated workbook | verify foreground color of cells after smart marker data source processing with Aspose.Cells
// Tags: smart markers conditional formatting validation Aspose.Cells | WorkbookDesigner data source processing C# | cell style foreground color comparison Aspose.Cells | Excel status column conditional formatting Aspose.Cells | conditional formatting color verification .NET | Aspose.Cells conditional formatting programmatic check

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ConditionalFormattingValidation
{
    // The example loads a template workbook that contains a smart marker for a "Status" column, fills a DataTable with status values, and processes the smart markers using WorkbookDesigner. It defines the expected foreground colors for each status (Approved‑green, Pending‑yellow, Rejected‑red) and iterates through the populated cells, retrieving each cell's style to compare the actual color against the expected one, logging any mismatches, and optionally saves the resulting workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                const string templatePath = "Template.xlsx";
                const string resultPath = "Result.xlsx";

                // Verify that the template file exists to avoid FileNotFoundException
                if (!File.Exists(templatePath))
                {
                    Console.WriteLine($"Template file \"{templatePath}\" not found.");
                    return;
                }

                // Load the template workbook that contains smart markers and conditional formatting
                Workbook workbook = new Workbook(templatePath);

                // Prepare data source for smart markers
                // Assume the template has a smart marker like "&=Status" in column B starting from row 2
                List<string> statusValues = new List<string>
                {
                    "Approved",
                    "Pending",
                    "Rejected",
                    "Approved",
                    "Pending"
                };

                // Create a DataTable with a single column "Status"
                DataTable dt = new DataTable("Data");
                dt.Columns.Add("Status", typeof(string));
                foreach (var val in statusValues)
                {
                    dt.Rows.Add(val);
                }

                // Process smart markers using WorkbookDesigner (correct API)
                try
                {
                    WorkbookDesigner designer = new WorkbookDesigner(workbook);
                    designer.SetDataSource(dt);
                    designer.Process();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Smart marker processing failed: {ex.Message}");
                    return;
                }

                // Define expected highlight colors for each status (ARGB format)
                var expectedColors = new Dictionary<string, Color>
                {
                    { "Approved", Color.FromArgb(0, 255, 0) },   // Green
                    { "Pending",  Color.FromArgb(255, 255, 0) }, // Yellow
                    { "Rejected", Color.FromArgb(255, 0, 0) }    // Red
                };

                // Validate conditional formatting by checking the actual cell background color
                Worksheet sheet = workbook.Worksheets[0];
                bool allValid = true;

                // Assuming status values are placed in column B (index 1) starting from row 2 (index 1)
                for (int i = 0; i < statusValues.Count; i++)
                {
                    int row = i + 1; // zero‑based index for row 2 onward
                    Cell cell = sheet.Cells[row, 1];
                    string status = cell.StringValue;

                    // Retrieve the cell's style after conditional formatting is applied
                    Style style = cell.GetStyle();
                    Color actualColor = style.ForegroundColor;

                    // Get the expected color for the current status
                    if (!expectedColors.TryGetValue(status, out Color expectedColor))
                    {
                        Console.WriteLine($"Unexpected status value '{status}' at cell {cell.Name}.");
                        allValid = false;
                        continue;
                    }

                    // Compare colors (ignoring alpha channel)
                    if (actualColor.ToArgb() != expectedColor.ToArgb())
                    {
                        Console.WriteLine($"Validation failed at {cell.Name}: status='{status}', expected color={expectedColor}, actual color={actualColor}.");
                        allValid = false;
                    }
                    else
                    {
                        Console.WriteLine($"Cell {cell.Name} correctly highlighted for status '{status}'.");
                    }
                }

                Console.WriteLine(allValid ? "All conditional formatting validations passed." : "Some conditional formatting validations failed.");

                // Save the workbook (optional, to inspect the result)
                try
                {
                    workbook.Save(resultPath);
                    Console.WriteLine($"Result workbook saved to \"{resultPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save result workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
