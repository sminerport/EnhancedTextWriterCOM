using System;
using System.Runtime.InteropServices;

namespace EnhancedTextWriterLib
{
   [ComVisible(true)]
   [InterfaceType(ComInterfaceType.InterfaceIsDual)]
   [Guid("97FBE95E-5EDC-494C-A703-39E230DB9A6C")]
   public interface IEnhancedTextWriter
   {
      void SaveToFile(string filePath, string text);
   }
}
