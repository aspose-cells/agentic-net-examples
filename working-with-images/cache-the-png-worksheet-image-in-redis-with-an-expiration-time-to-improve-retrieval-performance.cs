// Title: Cache a worksheet rendered as PNG in Redis with TTL using Aspose.Cells in C#
// AI Prompts: Write C# code that loads an Excel workbook, renders a specified worksheet to a PNG byte array with Aspose.Cells, and stores the byte array in Redis with a configurable expiration time, using reflection to avoid a compile‑time dependency on StackExchange.Redis. | Add error handling that falls back to writing the PNG file to a temporary folder when the Redis operation cannot be completed or the Redis assembly is missing.
// Common Searches: c# how to cache a rendered Excel worksheet as a PNG in Redis with expiration | aspnet store worksheet image in redis and set ttl using aspose.cells | fallback to local file when redis cache fails for image bytes in .net | using reflection to connect to StackExchange.Redis without adding a package | set expiration for binary data in redis from c#
// Tags: Aspose.Cells PNG rendering of worksheet | Redis binary cache with expiration | dynamic Redis connection via reflection | temporary file fallback for image cache | C# worksheet image caching pattern

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders a chosen worksheet to a PNG image in memory using Aspose.Cells, and attempts to cache the PNG bytes in Redis with a defined TTL. It uses reflection to invoke StackExchange.Redis methods, eliminating a hard compile‑time dependency. If Redis is unavailable or the assembly cannot be loaded, the image is saved to a temporary file as a graceful fallback.
public class WorksheetImageCache
{
    /// <param name="excelFilePath">Full path to the Excel file.</param>
    /// <param name="worksheetName">Name of the worksheet to render.</param>
    /// <param name="redisConnectionString">Redis connection string (e.g., "localhost:6379").</param>
    /// <param name="redisKey">Key under which the image would be stored in Redis.</param>
    /// <param name="expiration">Expiration time for the cached image.</param>
    public void CacheWorksheetImage(string excelFilePath, string worksheetName, string redisConnectionString, string redisKey, TimeSpan expiration)
    {
        try
        {
            // Verify that the Excel file exists to avoid FileNotFoundException
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException($"The Excel file was not found: {excelFilePath}");

            // Load the workbook from the file system
            var workbook = new Workbook(excelFilePath);

            // Retrieve the requested worksheet
            var worksheet = workbook.Worksheets[worksheetName];
            if (worksheet == null)
                throw new ArgumentException($"Worksheet '{worksheetName}' not found in the workbook.");

            // Configure image rendering options (default format is PNG)
            var imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Render the worksheet to a PNG image in memory
            byte[] imageBytes;
            using (var ms = new MemoryStream())
            {
                var sheetRender = new SheetRender(worksheet, imgOptions);
                // Render the first (and only) page because OnePagePerSheet = true
                sheetRender.ToImage(0, ms);
                imageBytes = ms.ToArray();
            }

            // Attempt to store the image in Redis; fallback to local file on any error
            try
            {
                // Check if StackExchange.Redis assembly is loaded
                var redisAssembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.GetName().Name.Equals("StackExchange.Redis", StringComparison.OrdinalIgnoreCase));

                if (redisAssembly != null)
                {
                    // Dynamically invoke Redis methods to avoid compile‑time dependency
                    var connectionMultiplexerType = redisAssembly.GetType("StackExchange.Redis.ConnectionMultiplexer");
                    var connectMethod = connectionMultiplexerType?.GetMethod("Connect", new[] { typeof(string) });
                    var redis = connectMethod?.Invoke(null, new object[] { redisConnectionString });

                    var getDatabaseMethod = redis?.GetType().GetMethod("GetDatabase", Type.EmptyTypes);
                    var db = getDatabaseMethod?.Invoke(redis, null);

                    // Resolve RedisKey and RedisValue types via reflection
                    var redisKeyType = redisAssembly.GetType("StackExchange.Redis.RedisKey");
                    var redisValueType = redisAssembly.GetType("StackExchange.Redis.RedisValue");

                    if (redisKeyType != null && redisValueType != null && db != null)
                    {
                        // Create instances of RedisKey and RedisValue
                        var redisKeyInstance = Activator.CreateInstance(redisKeyType, redisKey);
                        var redisValueInstance = Activator.CreateInstance(redisValueType, imageBytes);

                        // Get the StringSet method that accepts (RedisKey, RedisValue, TimeSpan?)
                        var stringSetMethod = db.GetType().GetMethod(
                            "StringSet",
                            new[] { redisKeyType, redisValueType, typeof(TimeSpan?) });

                        // Store the PNG bytes with the specified expiration
                        stringSetMethod?.Invoke(db, new object[] { redisKeyInstance, redisValueInstance, (TimeSpan?)expiration });
                    }
                    else
                    {
                        // Fallback: write image to a temporary file
                        WriteFallbackFile(redisKey, imageBytes);
                    }
                }
                else
                {
                    // Fallback: write image to a temporary file
                    WriteFallbackFile(redisKey, imageBytes);
                }
            }
            catch
            {
                // On any Redis‑related error, fallback to local file storage
                WriteFallbackFile(redisKey, imageBytes);
            }
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to preserve stack trace for the caller
            throw new ApplicationException("Failed to cache worksheet image.", ex);
        }
    }

    // Helper method to write the image to a temporary file
    private void WriteFallbackFile(string key, byte[] data)
    {
        try
        {
            var fallbackPath = Path.Combine(Path.GetTempPath(), $"{key}.png");
            File.WriteAllBytes(fallbackPath, data);
        }
        catch
        {
            // Swallow any exceptions from fallback storage to avoid breaking the main flow
        }
    }
}

// Minimal entry point to satisfy the project build
public static class Program
{
    public static void Main()
    {
        // Example usage (adjust paths and parameters as needed)
        // var cache = new WorksheetImageCache();
        // cache.CacheWorksheetImage("sample.xlsx", "Sheet1", "localhost:6379", "sheet1_image", TimeSpan.FromHours(1));
    }
}
