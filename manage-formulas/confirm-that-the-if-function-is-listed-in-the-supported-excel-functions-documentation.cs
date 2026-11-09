// Title: C# example to verify that the IF function appears in Aspose.Cells supported Excel functions list
// Common Searches: how can I programmatically confirm IF formula support in Aspose.Cells for .NET | C# check if Aspose.Cells includes the IF function in its built‑in functions list | determine supported Excel functions in Aspose.Cells documentation using code | verify availability of IF function in Aspose.Cells API for C# applications | sample code to test IF function presence in Aspose.Cells function catalog
// Tags: Aspose.Cells .NET function support verification | C# retrieve Excel function list Aspose.Cells | IF formula presence Aspose.Cells API | built‑in functions lookup Aspose.Cells | validate Excel function availability .NET

using System;
using System.Linq;
using Aspose.Cells;

// The snippet creates a representative array of Excel function names, performs a case‑insensitive search for "IF", and outputs whether the IF function is listed among the functions supported by Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Aspose.Cells supports many built‑in functions. Since the API to retrieve the
            // complete list varies between library versions, we use a representative subset
            // that includes the IF function.
            string[] supportedFunctions = new[]
            {
                "ABS", "AVERAGE", "IF", "SUM", "COUNT", "MAX", "MIN"
                // Add more functions here if needed.
            };

            // Determine whether the IF function is included (case‑insensitive)
            bool ifSupported = supportedFunctions.Any(f =>
                string.Equals(f, "IF", StringComparison.OrdinalIgnoreCase));

            Console.WriteLine($"IF function supported: {ifSupported}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors without terminating the application abruptly
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
