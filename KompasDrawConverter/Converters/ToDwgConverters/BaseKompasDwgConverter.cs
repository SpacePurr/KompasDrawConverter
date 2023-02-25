using Kompas6API5;

namespace KompasDrawConverter.Converters.ToDwgConverters
{
    public abstract class BaseKompasDwgConverter : BaseKompasConverter
    {
        protected BaseKompasDwgConverter(KompasObject kompas) : base(kompas)
        {
        }
        
        public abstract void ConvertToDwg(string filePath, string outputDirectory);
    }
}