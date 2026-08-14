// Title: C# – Retrieve Worksheet by Name, Show Formulas, and Save Workbook with Aspose.Cells
// Description: Loads an existing Excel file, accesses a worksheet using its name, enables the ShowFormulas flag so formulas appear instead of calculated values, and saves the workbook to a new file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# Excel manipulation | worksheet by name | ShowFormulas property | display formulas | save workbook | load Excel file | toggle formula view | Excel sheet debugging
// Common Searches: Aspose.Cells show formulas C# | retrieve worksheet by name Aspose.Cells | how to enable ShowFormulas for a sheet | save workbook after changing formula display | display formulas instead of values Aspose.Cells .NET
// Developer Intent: Open a workbook, locate a specific sheet via its name, turn on formula display for that sheet, and write the updated file back to disk.
// Use Cases: Auditor needs a copy of a financial model that reveals all underlying formulas. | Developer creates a debugging version of a spreadsheet to trace calculation errors. | Technical writer generates documentation that prints formulas rather than results. | QA team validates that cell references are correct by viewing formulas directly.
// AI Prompts: Generate C# code with Aspose.Cells that opens an Excel file, selects a worksheet by its name, sets ShowFormulas = true, and saves to a new file. | Explain the impact of the ShowFormulas property on worksheet rendering and how to apply it to multiple sheets in a workbook. | Provide a C# try‑catch example that handles a missing worksheet name when using Aspose.Cells. | Show how to toggle ShowFormulas for all worksheets in a workbook using a loop.

using System;
using Aspose.Cells;

// Loads an existing Excel file, accesses a worksheet using its name, enables the ShowFormulas flag so formulas appear instead of calculated values, and saves the workbook to a new file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the worksheet by its name (replace "Sheet1" with the actual sheet name)
        Worksheet worksheet = workbook.Worksheets["Sheet1"];

        // Set the worksheet to display formulas instead of calculated values
        worksheet.ShowFormulas = true;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
