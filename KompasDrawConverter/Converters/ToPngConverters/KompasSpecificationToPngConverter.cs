using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;

namespace KompasDrawConverter.Converters.ToPngConverters
{
    public class KompasSpecificationToPngConverter : BaseKompasPngConverter
    {
        public KompasSpecificationToPngConverter(KompasObject kompas) : base(kompas)
        {
        }

        public override void ConvertToRaster(string filePath, string outputDirectory)
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            
            var spec = (ksSpcDocument )Kompas.SpcDocument();
            spec.ksOpenDocument(filePath, 1);
            
            var imagePath = Path.Combine(outputDirectory, $"{fileNameWithoutExt}.png");
            
            var rasterFormatParam = (RasterFormatParam)spec.RasterFormatParam();
            InitRasterFormatParam(rasterFormatParam);
            
            spec.SaveAsToRasterFormat(imagePath, rasterFormatParam);
            spec.ksCloseDocument();

            Marshal.ReleaseComObject(rasterFormatParam);
            Marshal.ReleaseComObject(spec);
        }
    }
}