using System;
using System.Data;
using Aspose.Cells;

namespace ConditionalFormattingSmartMarkers
{
    class Program
    {
        static void Main()
        {
            // Load the template workbook that contains smart markers and conditional formatting.
            // The template file should have smart markers like "&=Data.Name" placed in cells
            // and a conditional formatting rule already defined on the target range.
            Workbook templateWorkbook = new Workbook("TemplateWithSmartMarkers.xlsx");

            // Prepare a data source that matches the smart marker name "Data".
            DataTable dataTable = new DataTable("Data");
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Score", typeof(double));

            // Sample rows – these values will replace the smart markers.
            dataTable.Rows.Add("Alice", 85.5);
            dataTable.Rows.Add("Bob", 72.0);
            dataTable.Rows.Add("Charlie", 91.2);

            // Set up the WorkbookDesigner with the loaded workbook.
            WorkbookDesigner designer = new WorkbookDesigner
            {
                Workbook = templateWorkbook
            };

            // Bind the data source to the smart marker name.
            designer.SetDataSource("Data", dataTable);

            // Process the smart markers. This will populate the cells with the data
            // while preserving any existing conditional formatting rules.
            designer.Process();

            // OPTIONAL: Verify that the conditional formatting rules are still present
            // and, for demonstration, adjust the priority of the first rule.
            Worksheet sheet = designer.Workbook.Worksheets[0];
            if (sheet.ConditionalFormattings.Count > 0)
            {
                // Get the first ConditionalFormatting collection.
                FormatConditionCollection conditions = sheet.ConditionalFormattings[0];

                // Iterate through all conditions in this collection.
                for (int i = 0; i < conditions.Count; i++)
                {
                    FormatCondition fc = conditions[i];
                    // Example: set higher priority (lower numeric value) for the first rule.
                    if (i == 0)
                    {
                        fc.Priority = 1; // highest priority
                    }
                }
            }

            // Save the processed workbook. The conditional formatting remains intact.
            designer.Workbook.Save("ProcessedOutput.xlsx");
        }
    }
}