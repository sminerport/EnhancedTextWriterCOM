using System;
using System.IO;
using System.Runtime.InteropServices;

namespace EnhancedTextWriterLib
{
   [ComVisible(true)]
   [ClassInterface(ClassInterfaceType.None)]
   [Guid("33FF6FC9-F461-4578-811D-04BB97BB9B31")]
   public class EnhancedTextWriter : IEnhancedTextWriter
   {
      public void SaveToFile(string filePath, string text)
      {
         string directoryPath = Path.GetDirectoryName(filePath);
         if (!Directory.Exists(directoryPath))
         {
            Directory.CreateDirectory(directoryPath);
         }
         File.WriteAllText(filePath, text);
      }
   }
}
