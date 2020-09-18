namespace PuppeteerSharp.Converter
{
  using System;
  using System.IO;
  using System.Linq;
  using System.Net;
  using System.Threading.Tasks;

  using PuppeteerSharp.Media;
  using Spire.Pdf;

  internal class Program
  {
    /// <summary>
    /// browser field
    /// </summary>
    private static Browser browser;

    /// <summary>
    /// fileDestination folder field
    /// </summary>
    private static string fileDestination = null;

    /// <summary>
    /// Main method
    /// </summary>
    private static void Main()
    {
      new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision).GetAwaiter().GetResult();
      browser = Puppeteer.LaunchAsync(new LaunchOptions
      {
        Headless = true,
        DefaultViewport = new ViewPortOptions()
        {
          IsMobile = false,
          IsLandscape = false,
          Height = 1080,
          Width = 1920
        }
      }).GetAwaiter().GetResult();

      Task.Run(async () => await LoadHtmlFile()).GetAwaiter().GetResult();
    }

    /// <summary>
    /// LoadHtmlFile method
    /// </summary>
    /// <returns></returns>
    private static async Task LoadHtmlFile()
    {
      var streamFileContent = await CreatePdfFileAsStream();
      if (streamFileContent != null)
      {
        try
        {
          CalculateFilePath("pdf");

          using (var fileStream = System.IO.File.Create(fileDestination))
          {
            streamFileContent.Seek(0, SeekOrigin.Begin);
            streamFileContent.CopyTo(fileStream);
          }

          CreateDocxFile();
        }
        catch { }
      }
    }

    /// <summary>
    /// CreateDocxFile method
    /// </summary>s
    private static void CreateDocxFile()
    {
      try
      {
        PdfDocument doc = new PdfDocument();
        doc.LoadFromFile(fileDestination);
        CalculateFilePath();
        doc.SaveToFile(fileDestination, FileFormat.DOCX);
      }
      catch { }
    }

    /// <summary>
    /// ReturnPath method
    /// </summary>
    /// <param name="folder"></param>
    /// <param name="extension"></param>
    /// <returns></returns>
    private static void CalculateFilePath(string extension = "docx")
    {
      var path = Environment.CurrentDirectory;
      fileDestination = Path.Combine(path, Guid.NewGuid() + $".{extension}");
    }

    /// <summary>
    /// WebClient ConvertHtmlToString method
    /// </summary>
    /// <returns></returns>
    private static string ConvertHtmlToString()
    {
      WebClient client = new WebClient();
      return client.DownloadString(CommonConstants.sourceFile);
    }

    #region "Converting methods"

    /// <summary>
    /// Pdf method, return pdf content as Task<Stream>
    /// </summary>
    /// <param name="page"></param>
    /// <returns></returns>
    private static async Task<Stream> CreatePdfFileAsStream()
    {
      try
      {
        var page = await browser.NewPageAsync().ConfigureAwait(true);
        //await page.GoToAsync("https://localhost:44351/Home/PrivacyPartial", WaitUntilNavigation.Networkidle0).ConfigureAwait(true);

        string htmlContent = ConvertHtmlToString();
        await page.SetContentAsync(htmlContent);

        Stream stream = await page.PdfStreamAsync(new PdfOptions()
        {
          Format = PaperFormat.A4,
          PreferCSSPageSize = true,
        }).ConfigureAwait(true);

        return stream;
      }
      catch { }

      return null;
    }

    /// <summary>
    /// ConvertHtmlToStream method
    /// </summary>
    /// <returns></returns>
    private static async Task<Stream> ConvertHtmlToStream()
    {
      MemoryStream ms = new MemoryStream();
      await using (FileStream file = new FileStream(CommonConstants.sourceFile, FileMode.Open, FileAccess.Read))
      {
        file.CopyTo(ms);
      }

      ms.Position = 0;
      return ms;
    }

    #endregion
  }
}
