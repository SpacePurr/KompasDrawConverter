using System.IO;

namespace KompasDrawConverter.Providers
{
    public class KompasToDwgConverterProvider : KompasFileConverter
    {
        private readonly string _outputDwgDirectory;
        
        public KompasToDwgConverterProvider(Options options) : base(options)
        {
            _outputDwgDirectory = Path.Combine(options.OutputDirectory, "dwg");
        }

        protected override void CreateDirectories()
        {
            if (!Directory.Exists(_outputDwgDirectory))
            {
                Directory.CreateDirectory(_outputDwgDirectory);
            }
        }

        protected override string[] GetAllowedFilesExtensions()
        {
            return new[] { "cdw", "spw" };
        }

        protected override void ConvertFile(KompasInstance kompasInstance, string filePath)
        {
            kompasInstance?.ConvertToDwg(filePath, _outputDwgDirectory);
        }
    }
}