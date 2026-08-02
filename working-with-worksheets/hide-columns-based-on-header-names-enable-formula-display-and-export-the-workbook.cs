// Title: C# Aspose.Cells Sample – Hide Columns by Header, Display Formulas, Save Workbook
// Description: A concise C# example that creates an Excel workbook with Aspose.Cells, hides any column whose header matches a predefined list (e.g., "Secret" and "Amount"), switches the worksheet to ShowFormulas mode, and writes the file to disk.
// Keywords: Aspose.Cells C# | .NET Excel automation | hide column by header | ShowFormulas property | export workbook | mask sensitive columns | Excel file generation | sample code GitHub | worksheet.HideColumn | formula view Excel
// Common Searches: Aspose.Cells hide column based on header name C# | How to display formulas instead of values with Aspose.Cells | Save Excel file after hiding sensitive columns using Aspose.Cells | C# example for worksheet.ShowFormulas Aspose.Cells | Hide multiple columns programmatically Aspose.Cells .NET
// Developer Intent: Hide specific columns, enable formula display, and export the workbook in one workflow.
// Use Cases: Prepare client‑ready reports that omit confidential fields such as passwords or amounts. | Debug complex spreadsheets by showing the underlying formulas in the exported file. | Automate Excel generation where certain columns must be invisible for compliance reasons.
// AI Prompts: Generate C# code with Aspose.Cells that hides columns whose header matches a given list and then saves the workbook. | Show how to set worksheet.ShowFormulas = true so that the exported Excel displays formulas, not results. | Provide a complete Aspose.Cells .NET snippet that creates a workbook, hides "Secret" and "Amount" columns, enables formula view, and writes Result.xlsx.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// A concise C# example that creates an Excel workbook with Aspose.Cells, hides any column whose header matches a predefined list (e.g., "Secret" and "Amount"), switches the worksheet to ShowFormulas mode, and writes the file to disk.
class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data with headers
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Secret");
        cells["D1"].PutValue("Amount");

        // Add some rows of data
        for (int row = 2; row <= 5; row++)
        {
            cells[$"A{row}"].PutValue(row - 1);
            cells[$"B{row}"].PutValue($"Item{row - 1}");
            cells[$"C{row}"].PutValue($"Hidden{row - 1}");
            cells[$"D{row}"].PutValue((row - 1) * 10);
        }

        // Insert a formula to demonstrate formula display
        cells["D6"].Formula = "SUM(D2:D5)";

        // Define header names of columns that should be hidden
        List<string> headersToHide = new List<string> { "Secret", "Amount" };

        // Hide columns whose header matches any name in the list
        int lastCol = cells.MaxColumn; // last column index containing data
        for (int col = 0; col <= lastCol; col++)
        {
            string header = cells[0, col].StringValue;
            if (headersToHide.Contains(header))
            {
                cells.HideColumn(col); // hide the column
            }
        }

        // Enable formula display (show formulas instead of calculated values)
        worksheet.ShowFormulas = true;

        // Export (save) the workbook to a file (lifecycle rule)
        workbook.Save("Result.xlsx");
    }
}
