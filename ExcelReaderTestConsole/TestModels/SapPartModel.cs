using ExcelReaderLibraryFW.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTestConsole.TestModels
{
  internal class SapPartModel
  {
    [ExcelField("Material")]
    public string PartNumber { get; set; }

    [ExcelField("Description")]
    public string Desc { get; set; }

    [ExcelField("Un-Restricted", 5)]
    public double Stock { get; set; }

    [ExcelField("Price")]
    public double Price { get; set; }

    public override string ToString()
    {
      return $"{PartNumber,-15} - {Desc,60} - {Stock} - {Price:C2}";
    }
  }
}
