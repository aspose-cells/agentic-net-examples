// Title: Programmatically verify if Aspose.Cells for .NET supports the VLOOKUP function by parsing the Supported Excel Functions page
// Common Searches: C# code to check if VLOOKUP is listed in Aspose.Cells supported functions documentation | how to programmatically confirm Excel formula support in Aspose.Cells .NET | retrieve Aspose.Cells supported functions list and search for VLOOKUP using HttpClient | determine if VLOOKUP works with Aspose.Cells by inspecting the supported functions page | sample C# script to validate VLOOKUP availability in Aspose.Cells library
// Tags: Aspose.Cells function availability query | C# HttpClient fetch Aspose.Cells docs | Excel formula support check .NET | VLOOKUP support detection Aspose.Cells | parse documentation for function list

using System;
using System.Net.Http;
using System.Threading.Tasks;

// The example uses HttpClient to download the Aspose.Cells Supported Excel Functions page, searches the HTML for the term "VLOOKUP" (case‑insensitive), and prints whether the function is listed as supported, handling any request errors that may occur.
class Program
{
    // URL of the Aspose.Cells Supported Excel Functions documentation
    private const string SupportedFunctionsUrl = "https://docs.aspose.com/cells/net/supported-excel-functions/";

    static async Task Main()
    {
        // Create an HttpClient instance (lifecycle managed by using statement)
        using (HttpClient client = new HttpClient())
        {
            try
            {
                // Load the documentation page
                string pageContent = await client.GetStringAsync(SupportedFunctionsUrl);

                // Verify if the VLOOKUP function is mentioned in the documentation
                bool isVlookupSupported = pageContent.IndexOf("VLOOKUP", StringComparison.OrdinalIgnoreCase) >= 0;

                // Output the verification result
                Console.WriteLine(isVlookupSupported
                    ? "VLOOKUP function is supported by Aspose.Cells."
                    : "VLOOKUP function is NOT listed in the supported functions documentation.");
            }
            catch (Exception ex)
            {
                // Handle any errors that occur during the request
                Console.WriteLine($"Error while accessing the documentation: {ex.Message}");
            }
        }
    }
}
