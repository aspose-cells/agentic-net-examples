// Title: Set MaxRowsOfSharedFormula to 100 in an Aspose.Cells workbook using C#
// AI Prompts: Configure the workbook to allow up to 100 rows in a shared formula block and save the file as XLSX with Aspose.Cells in C#. | Increase the shared formula row limit to 100 by modifying Workbook.Settings.MaxRowsOfSharedFormula and persist the workbook.
// Common Searches: aspnet set maxrowsofsharedformula 100 aspocells | c# enlarge shared formula capacity in Aspose.Cells workbook | how to raise shared formula rows count in Aspose.Cells .NET | Aspose.Cells workbook settings for shared formulas limit rows
// Tags: Aspose.Cells shared formula maximum rows | modify workbook settings to raise shared formula capacity C# | shared formula rows setting Aspose.Cells | increase shared formula rows with Aspose.Cells API

using System;
using Aspose.Cells;

// Creates a new Workbook, sets Workbook.Settings.MaxRowsOfSharedFormula to 100 to permit larger shared‑formula blocks, and saves the result as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Set the maximum number of rows that can be part of a shared formula block
        workbook.Settings.MaxRowsOfSharedFormula = 100;

        // Save the workbook to a file (you can change the format or path as needed)
        workbook.Save("Output.xlsx");
    }
}
