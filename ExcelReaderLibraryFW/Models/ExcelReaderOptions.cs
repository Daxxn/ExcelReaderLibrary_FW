using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderLibraryFW.Models
{
  /// <summary>
  /// Parser options used by the <see cref="ExcelReader"/> when reading files.
  /// </summary>
  public class ExcelReaderOptions
  {
    /// <summary>
    /// The index of the spreadsheet.
    /// </summary>
    public int WorkbookIndex { get; set; } = 0;

    /// <summary>
    /// The index of the column header. Usually the top of the sheet. (index 1)
    /// </summary>
    public int HeaderIndex { get; set; } = 1;

    /// <summary>
    /// The start of the data to be read. Usually after the header row.
    /// </summary>
    public int DataStartIndex { get; set; } = 2;

    /// <summary>
    /// The maximum number of columns to search.
    /// </summary>
    public int MaxColumnSize { get; set; } = 255;

    /// <summary>
    /// The column that will be checked to decide when to stop reading the file.
    /// <para/>
    /// EPPlus doesnt have an easy way to tell how may rows of data is contained in the spreadsheet.
    /// </summary>
    public int StopCheckColumn { get; set; } = 1;
    public string FileID { get; set; } = null;

    /// <summary>
    /// Initializes a new instance of the class.
    /// </summary>
    public ExcelReaderOptions() { }
  }
}
