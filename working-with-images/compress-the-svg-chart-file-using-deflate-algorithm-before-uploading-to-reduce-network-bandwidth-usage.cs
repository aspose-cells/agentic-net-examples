// Title: Compress an Aspose.Cells‑generated SVG chart with DeflateStream in C# before uploading via HttpClient
// Common Searches: how to use DeflateStream to compress SVG output from Aspose.Cells in a .NET console app | upload a deflate‑compressed SVG chart to a REST endpoint using HttpClient in C# | render an Excel worksheet chart to SVG with Aspose.Cells and reduce its size for network transfer | C# example for compressing Aspose.Cells SVG export with raw Deflate algorithm | best practice for sending compressed SVG payloads from Aspose.Cells to a web API
// Tags: Aspose.Cells render worksheet to SVG | DeflateStream compress SVG byte array | HttpClient post deflate encoded content | C# reduce SVG size for network bandwidth | raw Deflate compression without GZip header

using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgCompressionDemo
{
    // The sample creates a workbook with a column chart, renders the first worksheet to SVG using Aspose.Cells, compresses the SVG bytes with a raw DeflateStream (no GZip wrapper), and uploads the compressed payload to a specified HTTP endpoint via HttpClient with the appropriate deflate content‑encoding header.
    class Program
    {
        // Entry point – async to allow awaiting HttpClient calls
        static async Task Main(string[] args)
        {
            // Adjust these paths/URLs as needed for your environment
            const string outputSvgPath = "chart_output.svg";
            const string uploadUrl = "https://example.com/api/upload"; // <-- replace with real endpoint

            try
            {
                // 1️⃣ Create a workbook with sample data and a chart
                Workbook workbook = CreateSampleWorkbook();

                // 2️⃣ Render the first worksheet (containing the chart) to SVG in memory
                byte[] svgBytes = RenderWorksheetToSvg(workbook, 0);

                // Optionally write the raw SVG to disk for verification
                await File.WriteAllBytesAsync(outputSvgPath, svgBytes);
                Console.WriteLine($"SVG chart saved to '{outputSvgPath}' ({svgBytes.Length} bytes).");

                // 3️⃣ Compress the SVG using Deflate (raw, not GZip)
                byte[] compressedSvg = DeflateCompress(svgBytes);
                Console.WriteLine($"Compressed SVG size: {compressedSvg.Length} bytes.");

                // 4️⃣ Upload the compressed payload
                await UploadCompressedSvgAsync(compressedSvg, uploadUrl);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }

        private static Workbook CreateSampleWorkbook()
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("Apple");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("Banana");
            ws.Cells["B3"].PutValue(45);
            ws.Cells["A4"].PutValue("Cherry");
            ws.Cells["B4"].PutValue(25);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 8);
            Chart chart = ws.Charts[chartIdx];
            chart.Title.Text = "Fruit Sales";
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            return wb;
        }

        /// <summary>
        /// Renders the specified worksheet page to SVG and returns the raw bytes.
        /// </summary>
        private static byte[] RenderWorksheetToSvg(Workbook workbook, int sheetIndex)
        {
            Worksheet sheet = workbook.Worksheets[sheetIndex];

            // Configure SVG rendering options
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Svg,
                OnePagePerSheet = true,
                // Optional: make the SVG fit the viewport
                SVGFitToViewPort = true
            };

            // Render to a memory stream
            using (MemoryStream svgStream = new MemoryStream())
            {
                SheetRender renderer = new SheetRender(sheet, options);
                // Render the first (and only) page
                renderer.ToImage(0, svgStream);
                return svgStream.ToArray();
            }
        }

        /// <summary>
        /// Compresses the supplied byte array using the Deflate algorithm (raw, no header/footer).
        /// </summary>
        private static byte[] DeflateCompress(byte[] data)
        {
            using (MemoryStream output = new MemoryStream())
            {
                // LeaveOpen = true so we can read the underlying buffer after disposing DeflateStream
                using (DeflateStream deflate = new DeflateStream(output, CompressionLevel.Optimal, leaveOpen: true))
                {
                    deflate.Write(data, 0, data.Length);
                }
                // Reset position to read the compressed bytes
                output.Position = 0;
                return output.ToArray();
            }
        }

        /// <summary>
        /// Sends the compressed SVG payload to a remote HTTP endpoint.
        /// </summary>
        private static async Task UploadCompressedSvgAsync(byte[] compressedData, string url)
        {
            using (HttpClient client = new HttpClient())
            {
                // Prepare the content
                ByteArrayContent content = new ByteArrayContent(compressedData);
                content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                // Indicate that the payload is deflate‑compressed
                content.Headers.ContentEncoding.Add("deflate");

                Console.WriteLine($"Uploading compressed SVG to {url} ...");
                HttpResponseMessage response = await client.PostAsync(url, content);

                if (response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"Upload succeeded (status {(int)response.StatusCode}).");
                }
                else
                {
                    string responseBody = await response.Content.ReadAsStringAsync();
                    Console.Error.WriteLine($"Upload failed (status {(int)response.StatusCode}). Response: {responseBody}");
                }
            }
        }
    }
}
