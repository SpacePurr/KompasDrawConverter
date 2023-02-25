using Kompas6API5;

namespace KompasDrawConverter.Converters.ToPngConverters
{
    public abstract class BaseKompasPngConverter : BaseKompasConverter
    {
        protected BaseKompasPngConverter(KompasObject kompas) : base(kompas)
        {
        }
        
        public abstract void ConvertToRaster(string filePath, string outputDirectory);

        protected static void InitRasterFormatParam(RasterFormatParam rasterFormatParam)
        {
            rasterFormatParam.format = 2;
            rasterFormatParam.extResolution = 600;
            //rasterFormatParam.colorBPP = 1;
            //rasterFormatParam.colorType = 1;
            rasterFormatParam.greyScale = true;
        }
    }
}