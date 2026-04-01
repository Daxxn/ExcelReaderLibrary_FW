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
    /// The row where the header starts. If the header is only one line, this is the only needed setting.
    /// </summary>
    public int HeaderRowStartIndex { get; set; } = 1;

    /// <summary>
    /// The row where the header stops.
    /// <para/>
    /// Header is on one row = -1
    /// <para/>
    /// Header is on multiple rows = higher than <see cref="HeaderRowStartIndex"/>
    /// </summary>
    public int HeaderRowStopIndex { get; set; } = -1;

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

    /// <summary>
    /// Handle incosistencies and mising headers as errors.
    /// </summary>
    public bool Strict { get; set; } = false;

    /// <summary>
    /// When an objects property is a string, convert numbers found in the spreadsheet to strings.
    /// <para/>
    /// Otherwise the row will be skipped.
    /// </summary>
    public bool ConvertNumbersToString { get; set; } = false;

    /// <summary>
    /// Initializes a new instance of the class.
    /// </summary>
    public ExcelReaderOptions() { }
  }
}
