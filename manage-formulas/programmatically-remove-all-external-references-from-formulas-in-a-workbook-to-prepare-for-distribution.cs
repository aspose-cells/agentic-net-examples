// Title: Remove all external references from Excel formulas using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, calls the appropriate Aspose.Cells method to strip external links from all formulas, and saves the cleaned file. | Show how to use Aspose.Cells Workbook.RemoveExternalLinks to convert external references to #REF! errors in a C# application. | Explain the process of preparing an Excel file for distribution by programmatically clearing external formula links with Aspose.Cells in .NET.
// Common Searches: aspnet remove external links from Excel formulas Aspose.Cells | c# Aspose.Cells delete external references before sharing workbook | how to clear external workbook links using Aspose.Cells .NET API | remove external references from .xlsx programmatically with Aspose.Cells | Aspose.Cells Workbook.RemoveExternalLinks usage example
// Tags: external link removal Aspose.Cells | external formula reference cleanup .NET | strip external workbook links C# | prepare Excel file for distribution Aspose.Cells | replace external references with #REF! errors

using Aspose.Cells;
using System;

// // Loads an Excel workbook, removes all external links from its formulas (turning them into #REF! errors), and saves the cleaned file for safe distribution.
class Program
{
    static void Main()
    {
        // Load the workbook that contains external references
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all external references from formulas.
        // This clears external links and replaces them with #REF! errors.
        workbook.RemoveExternalLinks();

        // Save the workbook ready for distribution
        workbook.Save("output.xlsx");
    }
}
