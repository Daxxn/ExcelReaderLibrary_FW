using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderLibraryFW.Models
{
   public class ExcelReaderOptions
   {
      public int WorkbookIndex { get; set; } = 0;
      public int HeaderIndex { get; set; } = 0;
      public int DataIndex { get; set; } = 1;
      public int MaxColumnSize { get; set; } = 255;
      public string FileID { get; set; } = null;
      public ExcelReaderOptions() { }
   }
}
