using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;

namespace KompasDrawConverter.Converters.ToDwgConverters
{
    public class KompasDraw2DToDwgConverter : BaseKompasDwgConverter
    {
        public KompasDraw2DToDwgConverter(KompasObject kompas) : base(kompas)
        {
        }
        
        public override void ConvertToDwg(string filePath, string outputDirectory)
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            
            var doc2D = (ksDocument2D)Kompas.Document2D();
            doc2D.ksOpenDocument(filePath, true);
            
            var dwgPath = Path.Combine(outputDirectory, $"{fileNameWithoutExt}.dwg");

            doc2D.ksSaveToDXF(dwgPath);
            doc2D.ksCloseDocument();

            Marshal.ReleaseComObject(doc2D);
        }
    }
}