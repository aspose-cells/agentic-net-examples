using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Cells;

namespace AsposeCellsXmlConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Sample XML data
                string xml = @"<?xml version='1.0' encoding='UTF-8'?>
<Root>
    <Item>150</Item>
    <Item>80</Item>
    <Item>200</Item>
</Root>";

                // Parse XML and write values to cells (A1, A2, A3)
                XDocument doc = XDocument.Parse(xml);
                List<string> items = doc.Root?.Elements("Item").Select(e => e.Value).ToList() ?? new List<string>();

                for (int i = 0; i < items.Count; i++)
                {
                    cells[i, 0].PutValue(items[i]);
                }

                // Define the area that contains the imported values
                CellArea dataArea = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = items.Count - 1,
                    EndColumn = 0
                };

                // Add a conditional formatting collection to the worksheet
                int cfIndex = worksheet.ConditionalFormattings.Add();
                FormatConditionCollection cfCollection = worksheet.ConditionalFormattings[cfIndex];

                // Apply the conditional formatting to the data area
                cfCollection.AddArea(dataArea);

                // Add a condition: highlight cells with value > 100
                int conditionIdx = cfCollection.AddCondition(
                    FormatConditionType.CellValue,
                    OperatorType.GreaterThan,
                    "100",
                    null);

                // Set the style for the condition (red background, white bold font)
                FormatCondition condition = cfCollection[conditionIdx];
                condition.Style.BackgroundColor = Color.Red;
                condition.Style.Font.Color = Color.White;
                condition.Style.Font.IsBold = true;

                // Save the workbook
                string outputPath = "XmlConditionalFormattingOutput.xlsx";

                // Ensure the output directory exists (prevents FileNotFoundException on save)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}