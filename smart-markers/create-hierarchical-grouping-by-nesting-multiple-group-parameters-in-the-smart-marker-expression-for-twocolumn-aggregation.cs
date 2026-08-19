// Title: C# – Hierarchical grouping with multiple parameters in an Aspose.Cells pivot table (two‑column aggregation)
// Description: This example creates a workbook, adds Date, Category and Amount columns, builds a pivot table, places Date and Category as row fields and Amount as a data field, then groups the Date field by months and years and defines a custom group that merges categories "A" and "B" into "GroupAB" before refreshing and saving the file.
// Keywords: Aspose.Cells | C# pivot table grouping | hierarchical grouping | date months years | custom category group | two column aggregation | smart markers | .NET Excel automation | GitHub sample | pivot field GroupBy
// Common Searches: Aspose.Cells group date by month and year | Create custom group in Aspose.Cells pivot table | Nested row fields with hierarchical grouping C# | Smart marker expression two column aggregation | Aspose.Cells pivot table example GitHub
// Developer Intent: The developer wants to apply hierarchical grouping to a pivot table by using multiple group parameters for the Date field and a custom group for the Category field in a C# Aspose.Cells project.
// Use Cases: Financial reporting that summarizes amounts by year‑month while consolidating related categories. | Sales analysis where dates are drilled down from year to month and product categories are combined into custom groups. | Automated Excel generation for dashboards that require nested grouping of date and category dimensions.
// AI Prompts: Generate C# code with Aspose.Cells to group a pivot table Date field by both months and years and create a custom group for categories A and B. | Explain how to nest multiple group parameters in a smart marker expression for two‑column aggregation using Aspose.Cells. | Show how to retrieve pivot item indexes and define CustomPiovtFieldGroupItem to combine specific category values in a pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsSmartMarkerGrouping
{
    // This example creates a workbook, adds Date, Category and Amount columns, builds a pivot table, places Date and Category as row fields and Amount as a data field, then groups the Date field by months and years and defines a custom group that merges categories "A" and "B" into "GroupAB" before refreshing and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Populate sample data with two columns: Date and Category
            // ------------------------------------------------------------
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Amount");

            // Add several rows of data
            DateTime[] dates = {
                new DateTime(2023, 1, 5),
                new DateTime(2023, 1, 12),
                new DateTime(2023, 2, 3),
                new DateTime(2023, 2, 20),
                new DateTime(2023, 3, 15),
                new DateTime(2023, 3, 28)
            };
            string[] categories = { "A", "B", "A", "B", "A", "B" };
            double[] amounts = { 100, 150, 200, 250, 300, 350 };

            for (int i = 0; i < dates.Length; i++)
            {
                cells[i + 1, 0].PutValue(dates[i]);
                cells[i + 1, 1].PutValue(categories[i]);
                cells[i + 1, 2].PutValue(amounts[i]);
            }

            // ------------------------------------------------------------
            // Create a pivot table based on the data range
            // ------------------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("A1:C7", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add Date field to the row area (first level)
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");

            // Add Category field to the row area (second level, nested under Date)
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add Amount field to the data area (sum aggregation)
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // ------------------------------------------------------------
            // Hierarchical grouping:
            //   - Group the Date field by Months and Years (two group parameters)
            //   - Group the Category field by custom groups (e.g., A+B together)
            // ------------------------------------------------------------

            // Group the Date field (first row field) by Months and Years
            PivotField dateField = pivot.RowFields[0];
            DateTime startDate = new DateTime(2023, 1, 1);
            DateTime endDate   = new DateTime(2023, 12, 31);
            PivotGroupByType[] dateGroups = new PivotGroupByType[]
            {
                PivotGroupByType.Months,
                PivotGroupByType.Years
            };
            // Use interval = 1 and do not create a new field (group in place)
            dateField.GroupBy(startDate, endDate, dateGroups, 1, false);

            // Group the Category field (second row field) using a custom group:
            //   Create a custom group that combines "A" and "B" into "GroupAB"
            PivotField categoryField = pivot.RowFields[1];
            // Find the item indexes for "A" and "B"
            int indexA = -1, indexB = -1;
            for (int i = 0; i < categoryField.PivotItems.Count; i++)
            {
                string val = categoryField.PivotItems[i].Value?.ToString();
                if (val == "A") indexA = i;
                if (val == "B") indexB = i;
            }
            if (indexA != -1 && indexB != -1)
            {
                CustomPiovtFieldGroupItem[] customGroups = new CustomPiovtFieldGroupItem[]
                {
                    new CustomPiovtFieldGroupItem("GroupAB", new int[] { indexA, indexB })
                };
                // Create a new field for the custom group
                categoryField.GroupBy(customGroups, true);
            }

            // ------------------------------------------------------------
            // Refresh and calculate the pivot table to apply grouping
            // ------------------------------------------------------------
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            workbook.Save("HierarchicalGroupingSmartMarker.xlsx");
        }
    }
}
