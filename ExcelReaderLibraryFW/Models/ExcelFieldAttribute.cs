using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderLibraryFW.Models
{
  /// <summary>
  /// Excel column data for reading spreadsheets.
  /// </summary>
  [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
  public sealed class ExcelFieldAttribute : Attribute
  {
    private readonly string propName = null;
    private readonly bool ignoreCase = false;
    private readonly int columnStart = -1;
    private readonly int columnEnd = -1;

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// <para/>
    /// This will tell the reader to match the name of the property directly.
    /// </summary>
    public ExcelFieldAttribute() { }

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property in the spreadsheet.</param>
    public ExcelFieldAttribute(string propertyName)
    {
      propName = propertyName;
    }

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property in the spreadsheet.</param>
    /// <param name="ignoreCase">Ignore upper and lower case differences.</param>
    public ExcelFieldAttribute(string propertyName, bool ignoreCase)
    {
      propName = propertyName;
      this.ignoreCase = ignoreCase;
    }

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property in the spreadsheet.</param>
    /// <param name="columnStart">Ignores matches before this column in the spreadsheet.</param>
    public ExcelFieldAttribute(string propertyName, int columnStart)
    {
      propName = propertyName;
      this.columnStart = columnStart;
    }

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property in the spreadsheet.</param>
    /// <param name="columnStart">Ignores matches before this column in the spreadsheet.</param>
    /// <param name="columnEnd">Ignores matches after this column in the spreadsheet.</param>
    public ExcelFieldAttribute(string propertyName, int columnStart, int columnEnd)
    {
      propName = propertyName;
      this.columnStart = columnStart;
      this.columnEnd = columnEnd;
    }

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property in the spreadsheet.</param>
    /// <param name="columnStart">Ignores matches before this column in the spreadsheet.</param>
    /// <param name="ignoreCase">Ignore upper and lower case differences.</param>
    public ExcelFieldAttribute(string propertyName, int columnStart, bool ignoreCase)
    {
      propName = propertyName;
      this.columnStart = columnStart;
      this.ignoreCase = ignoreCase;
    }

    /// <summary>
    /// Mark a property for the <see cref="ExcelReader"/>.
    /// </summary>
    /// <param name="propertyName">The name of the property in the spreadsheet.</param>
    /// <param name="columnStart">Ignores matches before this column in the spreadsheet.</param>
    /// <param name="columnEnd">Ignores matches after this column in the spreadsheet.</param>
    /// <param name="ignoreCase">Ignore upper and lower case differences.</param>
    public ExcelFieldAttribute(string propertyName, int columnStart, int columnEnd, bool ignoreCase)
    {
      propName = propertyName;
      this.columnStart = columnStart;
      this.columnEnd = columnEnd;
      this.ignoreCase = ignoreCase;
    }

    /// <summary>
    /// Check the header value for a match with this property.
    /// </summary>
    /// <param name="input">The header input</param>
    /// <param name="propName">The name of the bound property</param>
    /// <param name="column">The current spreadsheet column</param>
    /// <returns>True if the column matches the property.</returns>
    public bool CheckProperty(string input, string propName, int column)
    {
      if (ColumnStart != -1)
      {
        if (column < ColumnStart) return false;
      }
      if (ColumnEnd != -1)
      {
        if (column > ColumnEnd) return false;
      }
      if (PropertyName == null && propName == input) return true;
      return ignoreCase ? PropertyName.ToLower() == input.ToLower() : PropertyName == input;
    }

    /// <summary>
    /// True to ignore upper and lower case differences.
    /// </summary>
    public bool IgnoreCase => ignoreCase;

    /// <summary>
    /// The name of the header in the spreadsheet.
    /// </summary>
    public string PropertyName => propName;

    /// <summary>
    /// The column to start checking for matches.
    /// </summary>
    public int ColumnStart => columnStart;

    /// <summary>
    /// The column to stop checking for matches.
    /// </summary>
    public int ColumnEnd => columnEnd;
  }
}
