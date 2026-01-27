using ExcelReaderLibraryFW.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderTestConsole.TestModels
{
  public class PicklistSagePartModel
  {
    [ExcelField("Product")]
    public string PartNumber { get; set; }

    [ExcelField("Open Qty")]
    public double OpenQuantity { get; set; }

    [ExcelField("Pick Qty")]
    public double PickQuantity { get; set; }

    [ExcelField("On Hand")]
    public double OnHandQuantity { get; set; }

    [ExcelField("Location")]
    public string Bin { get; set; }

    [ExcelField()]
    public string Description { get; set; }
  }
}
