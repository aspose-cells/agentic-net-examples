// Title: Create a Green Data Bar Conditional Format for a Named Range with Aspose.Cells (.NET)
// Description: Demonstrates how to generate a workbook, fill cells A1‑A10, define a named range "MyData", and apply a DataBar rule that uses automatic minimum/maximum values, solid green fill, and displays cell values. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# data bar | conditional formatting named range | automatic min max Excel bar | CellArea conditional format | DataBarFillType.Solid example | programmatic Excel data bar | Aspose.Cells workbook generation
// Common Searches: how to add a data bar to a named range using Aspose.Cells | Aspose.Cells conditional formatting with automatic scaling | C# create named range and apply data bar | set data bar color and show values in Aspose.Cells | retrieve named range for conditional formatting Aspose
// Developer Intent: Apply a DataBar rule to every cell in a defined range so each bar reflects the cell’s relative value within that range.
// Use Cases: Show performance metrics in column A with green bars that auto‑scale to the dataset. | Reuse a named block across multiple sheets and keep visual consistency with a single formatting rule. | Produce reports where bars update automatically when the underlying numbers change.
// AI Prompts: Write C# code with Aspose.Cells to add a red data bar to a named range "SalesData" using custom min and max thresholds. | Explain how to change the axis position and fill pattern of a DataBar applied to a named range in Aspose.Cells. | Provide steps to locate an existing named range and switch its data bar color to blue while preserving automatic scaling.

using System;
using System.Drawing;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDataBarNamedRange
{
    // Demonstrates how to generate a workbook, fill cells A1‑A10, define a named range "MyData", and apply a DataBar rule that uses automatic minimum/maximum values, solid green fill, and displays cell values. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column A (A1:A10)
                for (int i = 0; i < 10; i++)
                {
                    cells[i, 0].PutValue(i + 1); // Values 1..10
                }

                // Create a range object for A1:A10 and assign a name "MyData"
                AsposeRange dataRange = cells.CreateRange("A1", "A10");
                dataRange.Name = "MyData";

                // Retrieve the named range using the Name object
                Name namedRange = workbook.Worksheets.Names["MyData"];
                // Get the actual Range object (this will reflect the current address)
                AsposeRange range = namedRange.GetRange();

                // Add a new conditional formatting collection to the worksheet
                int cfIndex = sheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = sheet.ConditionalFormattings[cfIndex];

                // Define the area for the conditional formatting using the range's boundaries
                CellArea area = new CellArea
                {
                    StartRow = range.FirstRow,
                    EndRow = range.FirstRow + range.RowCount - 1,
                    StartColumn = range.FirstColumn,
                    EndColumn = range.FirstColumn + range.ColumnCount - 1
                };
                cfCollection.AddArea(area);

                // Add a DataBar condition
                int conditionIndex = cfCollection.AddCondition(FormatConditionType.DataBar);
                FormatCondition condition = cfCollection[conditionIndex];

                // Configure the DataBar properties
                DataBar dataBar = condition.DataBar;
                dataBar.MinCfvo.Type = FormatConditionValueType.AutomaticMin; // Minimum based on range
                dataBar.MaxCfvo.Type = FormatConditionValueType.AutomaticMax; // Maximum based on range
                dataBar.Color = Color.Green;                                   // Bar color
                dataBar.ShowValue = true;                                      // Show cell values
                dataBar.BarFillType = DataBarFillType.Solid;                   // Solid fill
                dataBar.AxisPosition = DataBarAxisPosition.Automatic;         // Default axis handling

                // Save the workbook
                string outputPath = "DataBar_NamedRange.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
