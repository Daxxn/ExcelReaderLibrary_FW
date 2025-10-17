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
    static string LocationsPath = @"C:\Users\Daxxn\Documents\WorkStuff\Cycle Counting\Locations_A101_10-17-25.xlsx";
    static bool ModernExcel { get; } = false;
    static void Main(string[] args)
    {
      ExcelPackage.License.SetNonCommercialPersonal("Daxxn Lantz");

      string path = LocationsPath;
      Console.WriteLine("Excel Reader Library Testing");
      Console.WriteLine($"Reading Excel File '{Path.GetFileName(path)}'");

      ExcelReaderOptions options = new ExcelReaderOptions()
      {
        DataStartIndex = 2,
        HeaderIndex = 1,
        StopCheckColumn = 1,
        WorkbookIndex = 0,
      };
      ExcelReader reader = new ExcelReader(options);

      var output = reader.Read<PartModel>(path);

      if (output == null)
      {
        Console.WriteLine("Output is null!");
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
          FileID = "ModernFile"
        };

      }
      else
      {
        testFile = @"F:\Code\C#\CSharpLibraries\ExcelReaderLibraryFW\TestOld.xls";
        options = new ExcelReaderOptions()
        {
          FileID = "OldFile"
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
