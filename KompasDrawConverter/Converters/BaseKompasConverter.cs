using Kompas6API5;

namespace KompasDrawConverter.Converters
{
    public abstract class BaseKompasConverter
    {
        protected KompasObject Kompas { get; }
        
        protected BaseKompasConverter(KompasObject kompas)
        {
            Kompas = kompas;
        }
    }
}