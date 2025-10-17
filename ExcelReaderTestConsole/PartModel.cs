using ExcelReaderLibraryFW.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTestConsole
{
  public enum Units
  {
    EA,
    IN,
    FT,
    RL,
  }

  public class PartModel
  {
    [ExcelField("Part Number")]
    public string PartNumber { get; set; }

    [ExcelField("Location")]
    public string BIN {  get; set; }

    [ExcelField("Description")]
    public string Desc { get; set; }

    [ExcelField]
    public Units Units { get; set; }

    [ExcelField]
    public double Stock { get; set; }

    public override string ToString()
    {
      return $"{PartNumber} {BIN} - {Stock} {Units} - {Desc}";
    }
  }
}
