// Title: Rename a named range and automatically update all formula references in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Change the defined name 'OldName' to 'NewName' and replace every occurrence of the old name in formulas throughout the workbook with Aspose.Cells. | Iterate over all worksheets and cells to detect formulas containing a specific named range and substitute the old range name with a new one using C# and Aspose.Cells. | Programmatically rename a named range, refresh dependent formulas in the Excel file, and save the updated workbook with Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells rename named range and update formulas in entire workbook | How to replace a defined name in all Excel formulas using Aspose.Cells .NET | Aspose.Cells iterate cells to modify formula references after renaming a range | Update Excel formula references after changing a named range with Aspose.Cells C#
// Tags: named range Text property Aspose.Cells | formula reference update Aspose.Cells | worksheet cell iteration Aspose.Cells | save workbook after modifications Aspose.Cells | Aspose.Cells defined name manipulation C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, locates the named range "OldName" in the NameCollection, changes its Text property to "NewName", then scans every worksheet and cell to replace occurrences of the old name in formulas before saving the modified workbook.
class RenameNamedRange
{
    static void Main()
    {
        const string inputFile = "input.xlsx";
        const string outputFile = "output.xlsx";
        const string oldName = "OldName";
        const string newName = "NewName";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file \"{inputFile}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputFile);

            // Access the collection of defined names (named ranges)
            NameCollection names = workbook.Worksheets.Names;

            // Find the named range with the old name
            Name definedName = names[oldName];
            if (definedName != null)
            {
                // Rename the defined name (use Text property)
                definedName.Text = newName;

                // Update all formula references that use the old name
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Cell cell in sheet.Cells)
                    {
                        if (cell.IsFormula && cell.Formula.Contains(oldName))
                        {
                            cell.Formula = cell.Formula.Replace(oldName, newName);
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine($"Warning: Named range \"{oldName}\" not found.");
            }

            // Save the workbook with the changes
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
