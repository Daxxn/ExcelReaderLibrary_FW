using ExcelReaderLibraryFW.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTestConsole.TestModels
{
  internal class SagePartModel
  {
    [ExcelField("Groups")]
    public string PartNumber { get; set; }

    [ExcelField("Description")]
    public string Desc { get; set; }

    [ExcelField("QTY")]
    public double Stock { get; set; }

    public override string ToString()
    {
      return $"{PartNumber} - {Desc,-20} - {Stock}";
    }
  }
}
