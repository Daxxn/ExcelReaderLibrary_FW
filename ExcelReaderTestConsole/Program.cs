using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelReaderLibraryFW;
using ExcelReaderLibraryFW.Models;
using OfficeOpenXml;

namespace ExcelReaderTestConsole
{
  internal class Program
  {
    static string RawSageLocationsPath = @"C:\Users\Daxxn\Documents\WorkStuff\Cycle Counting\RawLocations_A101_10-17-25.xlsx";
    static string SageReportPath = @"C:\Users\Daxxn\Documents\WorkStuff\VarianceAnalyzer\DailyReports\Inventory (Stock)_Stock by Site - Site 168 (Archive)_2025-08-25 090507_fe1ea1dc.xlsx";
    static string SapReportPath = @"C:\Users\Daxxn\Documents\WorkStuff\VarianceAnalyzer\DailyReports\ORWO 09-18-25XLSX.XLSX";
    static string LocationsPath = @"C:\Users\Daxxn\Documents\WorkStuff\Cycle Counting\Locations_A101_10-17-25.xlsx";
    static bool ModernExcel { get; } = false;
    static void Main(string[] args)
    {
      ExcelPackage.License.SetNonCommercialPersonal("Daxxn Lantz");

      //string path = SageReportPath;
      string path = SapReportPath;
      Console.WriteLine("Excel Reader Library Testing");
      Console.WriteLine($"Reading Excel File '{Path.GetFileName(path)}'");

      ExcelReaderOptions sageOptions = new ExcelReaderOptions()
      {
        HeaderRowStartIndex = 2,
        HeaderRowStopIndex = 3,
        DataStartIndex = 6,
        StopCheckColumn = 1,
        WorkbookIndex = 0,
      };
      ExcelReaderOptions sapOptions = new ExcelReaderOptions()
      {
        HeaderRowStartIndex = 1,
        DataStartIndex = 2,
        StopCheckColumn = 1,
      };
      //ExcelReader reader = new ExcelReader(sageOptions);
      ExcelReader reader = new ExcelReader(sapOptions);

      //var output = reader.Read<SagePartModel>(path);
      var output = reader.Read<SapPartModel>(path);

      if (output == null)
      {
        Console.WriteLine("Output is null!");
      }
      else if (output.Count() == 0)
      {
        Console.WriteLine("Output is empty.");
      }
      else
      {
        foreach (var item in output)
        {
          Console.WriteLine(item);
        }
      }

      Console.WriteLine("\nDONE.");
      Console.ReadLine();
    }

    private static void OldTest()
    {
      Console.WriteLine("Starting Excel Reader Test Console...");
      string testFile;
      ExcelReaderOptions options;
      if (ModernExcel)
      {
        testFile = @"F:\Code\C#\CSharpLibraries\ExcelReaderLibraryFW\ExcelParserTestDoc.xlsx";
        options = new ExcelReaderOptions()
        {
          //FileID = "ModernFile"
        };

      }
      else
      {
        testFile = @"F:\Code\C#\CSharpLibraries\ExcelReaderLibraryFW\TestOld.xls";
        options = new ExcelReaderOptions()
        {
          //FileID = "OldFile"
        };
      }
      ExcelReader reader = new ExcelReader(options);
      var models = reader.Read<TestModel>(testFile);

      foreach (var model in models)
      {
        Console.WriteLine(model);
      }

      Console.WriteLine("Done...");
      Console.ReadKey();
    }
  }
}
