// Title: Apply IF formulas and Icon Set conditional formatting to show category‑specific icons in an Excel file with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code using Aspose.Cells that creates a workbook, adds product and category columns, and inserts an IF formula in a hidden helper column to convert category names to numeric codes. | Write C# that adds an Icon Set conditional formatting to the helper column, assigning a distinct icon to each numeric code (1 = Electronics, 2 = Clothing, 3 = Food). | Provide C# to hide the helper column, configure the icon thresholds, and save the workbook as ProductCategoriesWithIcons.xlsx.
// Common Searches: aspnet aspose.cells conditional formatting icon set based on IF formula | c# how to display different icons for product categories in Excel using Aspose.Cells | map text category to numeric value with IF function in Aspose.Cells workbook | hide helper column after applying conditional formatting Aspose.Cells C# | set custom icon thresholds in Aspose.Cells conditional formatting C#
// Tags: conditional formatting icon set aspnet cells | if formula category mapping aspnet cells | suppress helper column visibility aspnet cells | custom icon thresholds aspnet cells | excel workbook category icons c#

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsConditionalImageDemo
{
    // Demonstrates creating an Excel workbook with product and category data, using IF formulas in a hidden helper column to map category text to numeric codes, applying an Icon Set conditional formatting that shows a different icon for each code, hiding the helper column, and saving the file as ProductCategoriesWithIcons.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Product name in column A, Category in column B
            // Category values: "Electronics", "Clothing", "Food"
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("Electronics");
            sheet.Cells["A3"].PutValue("T‑Shirt");
            sheet.Cells["B3"].PutValue("Clothing");
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue("Food");
            sheet.Cells["A5"].PutValue("Headphones");
            sheet.Cells["B5"].PutValue("Electronics");
            sheet.Cells["A6"].PutValue("Jeans");
            sheet.Cells["B6"].PutValue("Clothing");
            sheet.Cells["A7"].PutValue("Bread");
            sheet.Cells["B7"].PutValue("Food");

            // Hidden helper column C will contain a numeric code derived from the category
            // Using IF formulas to map text categories to numbers:
            // Electronics -> 1, Clothing -> 2, Food -> 3
            for (int row = 1; row <= 7; row++)
            {
                // Formula placed in C(row+1) because rows are zero‑based internally
                string formula = $"IF(B{row + 1}=\"Electronics\",1,IF(B{row + 1}=\"Clothing\",2,3))";
                sheet.Cells[row, 2].Formula = formula; // column index 2 = C
            }

            // Apply an Icon Set conditional formatting to the helper column (C2:C8)
            int cfIndex = sheet.ConditionalFormattings.Add();
            FormatConditionCollection fcc = sheet.ConditionalFormattings[cfIndex];

            // Define the range for the conditional formatting
            CellArea area = new CellArea
            {
                StartRow = 1,   // C2
                EndRow = 7,     // C8
                StartColumn = 2,
                EndColumn = 2
            };
            fcc.AddArea(area);

            // Add an IconSet condition
            int conditionIdx = fcc.AddCondition(FormatConditionType.IconSet);
            FormatCondition iconCondition = fcc[conditionIdx];

            // Use a custom icon set (three icons)
            // Add three different icons – you can replace the IconSetType values with any icons you prefer
            iconCondition.IconSet.CfIcons.Add(IconSetType.Arrows3, 0);      // Icon for Electronics
            iconCondition.IconSet.CfIcons.Add(IconSetType.ArrowsGray3, 1); // Icon for Clothing
            iconCondition.IconSet.CfIcons.Add(IconSetType.Boxes5, 2);      // Icon for Food

            // Set the thresholds so that:
            // 1 → first icon, 2 → second icon, 3 → third icon
            // The thresholds are inclusive; we use numeric values directly.
            iconCondition.SetFormula1("0", false, false); // lower bound (not used but required)
            iconCondition.SetFormula2("3", false, false); // upper bound
            iconCondition.Operator = OperatorType.Between;

            // Hide the helper column so only the icons are visible to the user
            sheet.Cells.HideColumn(2); // column index 2 = C

            // Save the workbook
            workbook.Save("ProductCategoriesWithIcons.xlsx");
        }
    }
}
