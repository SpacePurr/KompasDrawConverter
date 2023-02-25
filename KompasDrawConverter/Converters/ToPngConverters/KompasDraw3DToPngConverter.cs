using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;

namespace KompasDrawConverter.Converters.ToPngConverters
{
    public class KompasDraw3DToPngConverter : BaseKompasPngConverter
    {
        public KompasDraw3DToPngConverter(KompasObject kompas) : base(kompas)
        {
        }

        public override void ConvertToRaster(string filePath, string outputDirectory)
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            
            var doc3D = (ksDocument3D)Kompas.Document3D();
            doc3D.Open(filePath);
            
            var imagePath = Path.Combine(outputDirectory, $"{fileNameWithoutExt}.png");

            var rasterFormatParam = (RasterFormatParam)doc3D.RasterFormatParam();
            InitRasterFormatParam(rasterFormatParam);
            
            doc3D.SaveAsToRasterFormat(imagePath, rasterFormatParam);
            doc3D.close();

            Marshal.ReleaseComObject(rasterFormatParam);
            Marshal.ReleaseComObject(doc3D);
        }
    }
}