// Title: Insert VAR subtotals in column I with outline view using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that adds variance (VAR) subtotals to column I of an Aspose.Cells worksheet and enables outline grouping for hierarchical display. | Create a version‑agnostic reflection call to Worksheet.Subtotal that applies the VAR function and configures replace, page break, summary below, and summary right options.
// Common Searches: how to add variance subtotal to a specific column with Aspose.Cells C# | Aspose.Cells Subtotal method reflection example for VAR function | enable outline grouping for subtotals in Excel using Aspose.Cells .NET | C# code to calculate VAR subtotals and show them as collapsed groups | using Worksheet.Subtotal to insert VAR subtotals on column I in Aspose.Cells
// Tags: VAR subtotal Aspose.Cells C# | outline grouping subtotals .NET | reflection Worksheet.Subtotal call | subtotal column I Excel Aspose | variance aggregation workbook

using System;
using System.IO;
using System.Reflection;
using Aspose.Cells;

// The program creates a workbook, fills column I with numeric values, then uses reflection to locate and invoke Worksheet.Subtotal with the VAR function, configuring replace, page break, summary below, and outline grouping options, and finally saves the file as SubtotalsVar.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];

            // Header
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["I1"].PutValue("Value");

            // Sample data rows
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["I2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("A");
            worksheet.Cells["I3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("B");
            worksheet.Cells["I4"].PutValue(30);
            worksheet.Cells["A5"].PutValue("B");
            worksheet.Cells["I5"].PutValue(40);

            // Add subtotals on column I (zero‑based index 8) using the VAR function.
            // Use reflection to stay compatible with different Aspose.Cells versions.
            try
            {
                // Get the Subtotal method via reflection
                MethodInfo subtotalMethod = typeof(Worksheet).GetMethod(
                    "Subtotal",
                    BindingFlags.Instance | BindingFlags.Public,
                    null,
                    new Type[]
                    {
                        typeof(int),                     // column
                        typeof(Enum),                    // function (enum type, use Enum as placeholder)
                        typeof(bool),                    // replace
                        typeof(bool),                    // pageBreak
                        typeof(bool),                    // summaryBelow
                        typeof(bool),                    // summaryRight
                        typeof(int[])                    // totalColumns
                    },
                    null);

                // If the exact signature is not found (e.g., older version), try without the enum placeholder
                if (subtotalMethod == null)
                {
                    subtotalMethod = typeof(Worksheet).GetMethod(
                        "Subtotal",
                        BindingFlags.Instance | BindingFlags.Public);
                }

                if (subtotalMethod != null)
                {
                    // Resolve the SubtotalFunction enum type
                    Type enumType = typeof(Worksheet).Assembly.GetType("Aspose.Cells.SubtotalFunction");
                    object functionValue;

                    if (enumType != null && Enum.IsDefined(enumType, "Var"))
                    {
                        // Use the enum value for VAR
                        functionValue = Enum.Parse(enumType, "Var");
                    }
                    else
                    {
                        // Fallback to the integer value for VAR (10)
                        functionValue = 10;
                    }

                    // Prepare parameters for the Subtotal method
                    object[] parameters = new object[]
                    {
                        8,                 // column index (I)
                        functionValue,     // function (VAR)
                        true,              // replace existing subtotals
                        false,             // page break
                        true,              // summary below
                        false,             // summary right
                        new int[] { 8 }    // total columns
                    };

                    // Invoke the method
                    subtotalMethod.Invoke(worksheet, parameters);
                }
                else
                {
                    Console.WriteLine("Subtotal method is not available in this Aspose.Cells version.");
                }
            }
            catch (TargetInvocationException tie)
            {
                Console.WriteLine($"Error invoking Subtotal: {tie.InnerException?.Message ?? tie.Message}");
            }

            // Save the workbook
            string outputPath = "SubtotalsVar.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
