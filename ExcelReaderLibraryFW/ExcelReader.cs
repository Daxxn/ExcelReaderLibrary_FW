using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;


using ExcelReaderLibraryFW.Models;
using OfficeOpenXml;

namespace ExcelReaderLibraryFW
{
  /// <summary>
  /// Reads an Excel file and builds the data.
  /// </summary>
  public class ExcelReader
  {
    #region Local Props
    private const int MaxTimeout = 100;
    /// <summary>
    /// Options used when reading spreadsheets.
    /// </summary>
    public ExcelReaderOptions Options { get; private set; } = new ExcelReaderOptions();

    private Dictionary<int, PropertyInfo> PropertyHeaders { get; set; } = new Dictionary<int, PropertyInfo>();
    #endregion

    #region Constructors
    /// <summary>
    /// Creates a new instance of <see cref="ExcelReader"/> with the required options.
    /// </summary>
    /// <param name="options">The options used when reading the spreadsheet.</param>
    public ExcelReader(ExcelReaderOptions options)
    {
      Options = options;
    }
    #endregion

    #region Methods
    /// <summary>
    /// Read Excel spreadsheet and keep track of errors.
    /// </summary>
    /// <typeparam name="T">The model of the data.</typeparam>
    /// <param name="filePath">The path to the file.</param>
    /// <returns>A tuple containing the read data and any errors.</returns>
    /// <exception cref="Exception"></exception>
    public (IEnumerable<T>, List<Exception>) ReadVerbose<T>(string filePath) where T : class, new()
    {
      if (File.Exists(filePath))
      {
        var ext = Path.GetExtension(filePath).ToLower();
        if (ext == ".xlsx" || ext == ".xls")
        {
          if (Options.HeaderRowStartIndex >= Options.DataStartIndex)
          {
            throw new Exception("The header index cannot be greater than or equal to the data index");
          }
          using (var package = new ExcelPackage(filePath))
          {
            var sheet = package.Workbook.Worksheets[Options.WorkbookIndex];

            try
            {
              ParseHeader<T>(sheet);
            }
            catch (Exception)
            {
              throw;
            }

            List<T> data = new List<T>();
            List<Exception> errors = new List<Exception>();
            var rowEnd = sheet.Rows.EndRow;
            for (int r = Options.DataStartIndex; r < rowEnd; r++)
            {
              try
              {
                if (sheet.Cells[r, Options.StopCheckColumn].Value is null)
                {
                  break;
                }
                data.Add(ParseDataRow<T>(sheet, r));
              }
              catch (Exception e)
              {
                errors.Add(e);
              }
            }
            return (data, errors);
          }
        }
        else
        {
          throw new Exception("File isnt a valid type. Needs to be either \".xls\" or \".xlsx\"");
        }
      }
      else
      {
        throw new Exception("File cannot be found.");
      }
    }

    /// <summary>
    /// Read Excel spreadsheet
    /// </summary>
    /// <typeparam name="T">The model to create objects from.</typeparam>
    /// <param name="filePath">The path to the excel file.</param>
    /// <returns>A list of parsed objects from the spreadsheet</returns>
    /// <exception cref="Exception"></exception>
    public IEnumerable<T> Read<T>(string filePath) where T : class, new()
    {
      if (File.Exists(filePath))
      {
        var ext = Path.GetExtension(filePath).ToLower();
        if (ext == ".xlsx" || ext == ".xls")
        {
          if (Options.HeaderRowStartIndex >= Options.DataStartIndex)
          {
            throw new Exception("The header index cannot be greater than or equal to the data index");
          }
          using (var package = new ExcelPackage(filePath))
          {
            var sheet = package.Workbook.Worksheets[Options.WorkbookIndex];

            ParseHeader<T>(sheet);
            List<T> data = new List<T>();
            var rowEnd = sheet.Rows.EndRow;
            for (int r = Options.DataStartIndex; r < rowEnd; r++)
            {
              if (sheet.Cells[r, Options.StopCheckColumn].Value is null)
              {
                break;
              }
              data.Add(ParseDataRow<T>(sheet, r));
            }
            return data;
          }
        }
        else
        {
          throw new Exception("File isnt a valid type. Needs to be either \".xls\" or \".xlsx\"");
        }
      }
      else
      {
        throw new Exception("File cannot be found.");
      }
    }

    private void ParseHeader<T>(ExcelWorksheet sheet) where T : class, new()
    {
      if (Options.HeaderRowStopIndex != -1)
      {
        ParseMultiRowHeader<T>(sheet);
      }
      else
      {
        ParseSingleRowHeader<T>(sheet);
      }
    }

    private void ParseMultiRowHeader<T>(ExcelWorksheet sheet) where T : class, new()
    {
      PropertyHeaders.Clear();
      var props = new Queue<PropertyInfo>(new T().GetType().GetProperties());
      var start = sheet.Columns.StartColumn;
      var end = sheet.Columns.EndColumn;
      while (props.Count > 0)
      {
        var prop = props.Dequeue();
        bool found = false;
        var excelFields = prop.GetCustomAttributes<ExcelFieldAttribute>();
        if (excelFields.Count() == 0) continue;

        for (int r = Options.HeaderRowStartIndex; r <= Options.HeaderRowStopIndex; r++)
        {
          for (int c = start; c <= end; c++)
          {
            if (PropertyHeaders.ContainsKey(c)) continue;

            if (sheet.Cells[r, c].Value is string headerName)
            {
              if (string.IsNullOrEmpty(headerName)) continue;

              if (excelFields.Any())
              {
                if (excelFields.Any(field => field.CheckProperty(headerName, prop.Name, c)))
                {
                  if (!PropertyHeaders.ContainsKey(c))
                  {
                    PropertyHeaders.Add(c, prop);
                    found = true;
                    break;
                  }
                }
              }
            }
          }

          if (found) break;
        }
      }
    }

    private void ParseSingleRowHeader<T>(ExcelWorksheet sheet) where T : class, new()
    {
      PropertyHeaders.Clear();
      var props = new Queue<PropertyInfo>(new T().GetType().GetProperties());
      var start = sheet.Columns.StartColumn;
      var end = sheet.Columns.EndColumn;
      while (props.Count != 0)
      {
        var prop = props.Dequeue();

        for (int c = start; c < end + 1; c++)
        {
          if (PropertyHeaders.ContainsKey(c)) continue;

          if (sheet.Cells[Options.HeaderRowStartIndex, c].Value is string headerName)
          {
            if (string.IsNullOrEmpty(headerName)) continue;

            var excelFields = prop.GetCustomAttributes<ExcelFieldAttribute>();
            if (excelFields.Any())
            {
              if (excelFields.Any(field => field.CheckProperty(headerName, prop.Name, c)))
              {
                if (!PropertyHeaders.ContainsKey(c))
                {
                  PropertyHeaders.Add(c, prop);
                  break;
                }
              }
            }
          }
        }
      }
    }

    private T ParseDataRow<T>(ExcelWorksheet sheet, int rowCount) where T : class, new()
    {
      T newObj = new T();
      var start = sheet.Columns.StartColumn;
      var end = sheet.Columns.EndColumn;
      for (int i = start; i < end + 1; i++)
      {
        if (PropertyHeaders.ContainsKey(i))
        {
          var prop = PropertyHeaders[i];
          ParseProperty(sheet.Cells[rowCount, i].Value, prop, newObj);
        }
      }
      return newObj;
    }

    //private void SetReaderIndex(IExcelDataReader reader, bool header = false)
    //{
    //  reader.Reset();
    //  if (Options.WorkbookIndex != 0)
    //  {
    //    if (reader.ResultsCount >= Options.WorkbookIndex)
    //    {
    //      for (int i = 0; i < Options.WorkbookIndex; i++)
    //      {
    //        reader.NextResult();
    //      }
    //    }
    //  }
    //  if (header)
    //  {
    //    if (Options.HeaderIndex != 0)
    //    {
    //      for (int i = 0; i < Options.HeaderIndex; i++)
    //      {
    //        reader.Read();
    //      }
    //    }
    //  }
    //  else
    //  {
    //    for (int i = 0; i < Options.DataIndex; i++)
    //    {
    //      reader.Read();
    //    }
    //  }
    //}

    //private void ParseProperty<T>(object value, PropertyInfo prop, T newObj, bool ignoreCase = false) where T : class, new()
    //{
    //  if (value == null) return;
    //  if (prop.PropertyType == value.GetType())
    //  {
    //    prop.SetValue(newObj, value);
    //  }
    //  else if (prop.PropertyType.Name == "string")
    //  {
    //    prop.SetValue(newObj, value.ToString());
    //  }
    //  else if (prop.PropertyType.Name == "int")
    //  {
    //    prop.SetValue(newObj, value);
    //  }
    //}

    private void ParseProperty<T>(object value, PropertyInfo prop, T newObj, bool ignoreCase = false) where T : class, new()
    {
      if (value == null) return;
      if (prop.PropertyType == value.GetType())
      {
        prop.SetValue(newObj, value);
      }
      else if (value is double)
      {
        if (prop.PropertyType == typeof(int))
        {
          prop.SetValue(newObj, Convert.ToInt32(value));
        }
        else if (prop.PropertyType == typeof(float))
        {
          prop.SetValue(newObj, Convert.ToSingle(value));
        }
        else if (prop.PropertyType == typeof(decimal))
        {
          prop.SetValue(newObj, Convert.ToDecimal(value));
        }
        else if (prop.PropertyType == typeof(byte))
        {
          prop.SetValue(newObj, Convert.ToByte(value));
        }
      }
      else if (value is string v)
      {
        if (prop.PropertyType == typeof(char))
        {
          prop.SetValue(newObj, Convert.ToChar(value));
        }
        else if (prop.PropertyType.IsEnum)
        {
          try
          {
            prop.SetValue(newObj, Enum.Parse(prop.PropertyType, v, ignoreCase));
          }
          catch (Exception)
          {

          }
        }
      }
    }
    #endregion

    #region Full Props

    #endregion
  }
}
