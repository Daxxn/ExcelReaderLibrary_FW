using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderLibraryFW.Models
{
   [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = true)]
   public sealed class ExcelFieldAttribute : Attribute
   {
      private readonly string propName = null;
      private readonly string fileID = null;
      private readonly bool ignoreCase = false;

      // This is a positional argument
      public ExcelFieldAttribute() { }
      public ExcelFieldAttribute(string propertyName)
      {
         propName = propertyName;
      }
      public ExcelFieldAttribute(string propertyName, bool ignoreCase)
      {
         propName = propertyName;
         this.ignoreCase = ignoreCase;
      }
      public ExcelFieldAttribute(string propertyName, string fileTypeID)
      {
         propName = propertyName;
         fileID = fileTypeID;
      }
      public ExcelFieldAttribute(string propertyName, string fileTypeID, bool ignoreCase)
      {
         propName = propertyName;
         fileID = fileTypeID;
         this.ignoreCase = ignoreCase;
      }

      public bool CheckProperty(string input, string propName, string fileID)
      {
         if (FileID != null)
         {
            if (fileID != FileID) return false;
         }
         if (PropertyName == null && propName == input) return true;
         return ignoreCase ? PropertyName.ToLower() == input.ToLower() : PropertyName == input;
      }

      public bool IgnoreCase => ignoreCase;
      public string PropertyName => propName;
      public string FileID => fileID;
   }
}
