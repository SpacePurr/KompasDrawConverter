using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;
using Kompas6Constants;
using KompasAPI7;

namespace KompasDrawConverter.Converters.ToPngConverters
{
    public class KompasDraw2DToPngConverter : BaseKompasPngConverter
    {
        public KompasDraw2DToPngConverter(KompasObject kompas) : base(kompas)
        {
        }

        public override void ConvertToRaster(string filePath, string outputDirectory)
        {
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            
            //var kompas7 = (_Application)Kompas.ksGetApplication7();
            
            //IKompasDocument kompasDocument = kompas7.ActiveDocument;
            
            var doc2D = (ksDocument2D)Kompas.Document2D();
            doc2D.ksOpenDocument(filePath, true);
            
            var imagePath = Path.Combine(outputDirectory, $"{fileNameWithoutExt}.png");
            
            var rasterFormatParam = (RasterFormatParam)doc2D.RasterFormatParam();
            InitRasterFormatParam(rasterFormatParam);

            doc2D.SaveAsToRasterFormat(imagePath, rasterFormatParam);
            doc2D.ksCloseDocument();

            //kompasDocument.Close(DocumentCloseOptions.kdDoNotSaveChanges);
            
            Marshal.ReleaseComObject(rasterFormatParam);
            Marshal.ReleaseComObject(doc2D);
        }
    }
}