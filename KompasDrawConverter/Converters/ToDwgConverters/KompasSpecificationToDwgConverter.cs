using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;

namespace KompasDrawConverter.Converters.ToDwgConverters
{
    public class KompasSpecificationToDwgConverter : BaseKompasDwgConverter
    {
        public KompasSpecificationToDwgConverter(KompasObject kompas) : base(kompas)
        {
            
        }
        
        public override void ConvertToDwg(string filePath, string outputDirectory)
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            
            var spec = (ksSpcDocument )Kompas.SpcDocument();
            spec.ksOpenDocument(filePath, 1);
            
            var dwgPath = Path.Combine(outputDirectory, $"{fileNameWithoutExt}.dwg");

            spec.ksSaveToDXF(dwgPath);
            spec.ksCloseDocument();

            Marshal.ReleaseComObject(spec);
        }
    }
}