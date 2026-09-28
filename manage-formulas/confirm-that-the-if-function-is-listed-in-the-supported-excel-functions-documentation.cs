// Title: Check if the IF Excel function is supported in Aspose.Cells using C# reflection of Workbook.GetSupportedFunctions
// AI Prompts: Write a C# program that uses reflection to invoke Workbook.GetSupportedFunctions and returns true when the IF function appears in the returned list. | Generate C# code that gracefully handles the absence of Workbook.GetSupportedFunctions and still reports whether a specified Excel formula (e.g., IF) is available in the current Aspose.Cells version.
// Common Searches: how to verify IF function support in Aspose.Cells .NET | using reflection to list supported Excel functions in Aspose.Cells C# | determine if a specific Excel formula is available at runtime with Aspose.Cells | fallback when Workbook.GetSupportedFunctions method is missing in Aspose.Cells
// Tags: Aspose.Cells runtime function support check | C# reflection GetSupportedFunctions | verify Excel IF formula availability Aspose.Cells | handle missing GetSupportedFunctions method | enumerate supported Excel functions Aspose.Cells

using System;
using System.Linq;
using System.Reflection;
using Aspose.Cells;

// The example uses C# reflection to call the static Workbook.GetSupportedFunctions method (if present), retrieves the array of supported Excel function names, and case‑insensitively checks whether the IF function is included, outputting the support status while handling the method's possible absence.
class Program
{
    static void Main()
    {
        try
        {
            // Attempt to retrieve the list of all supported Excel functions via reflection.
            // This avoids compile‑time dependency on a specific Aspose.Cells version.
            string[] supportedFunctions = Array.Empty<string>();
            MethodInfo getFuncsMethod = typeof(Workbook).GetMethod(
                "GetSupportedFunctions",
                BindingFlags.Public | BindingFlags.Static);

            if (getFuncsMethod != null)
            {
                // Invoke the static method if it exists.
                object result = getFuncsMethod.Invoke(null, null);
                supportedFunctions = result as string[] ?? Array.Empty<string>();
            }
            else
            {
                Console.WriteLine("The method Workbook.GetSupportedFunctions() is not available in the current Aspose.Cells version.");
            }

            // Check if the "IF" function is present (case‑insensitive).
            bool ifSupported = supportedFunctions.Any(f =>
                string.Equals(f, "IF", StringComparison.OrdinalIgnoreCase));

            // Output the result.
            Console.WriteLine($"Is the IF function supported? {ifSupported}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
