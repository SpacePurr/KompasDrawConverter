using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using Kompas6API5;
using KompasAPI7;
using KompasDrawConverter.Converters;
using KompasDrawConverter.Converters.ToDwgConverters;
using KompasDrawConverter.Converters.ToPngConverters;
using Spectre.Console;

namespace KompasDrawConverter
{
    public class KompasInstance : IDisposable
    {
        private static KompasInstance _instance;
        
        private KompasObject _kompas;
        
        private KompasInstance()
        {
            AnsiConsole.Status()
            .Start("Запуск Компас-3D...", ctx => 
            {
                var t = Type.GetTypeFromProgID("KOMPAS.Application.5");
                _kompas = (KompasObject)Activator.CreateInstance(t);
            });
        }
        
        public static KompasInstance GetInstance()
        {
            return _instance ??= new KompasInstance();
        }
    
        public string ConvertToRaster(string filePath, string outputDirectory)
        {
            var fileType = GetKompasFileType(filePath);
            
            var fileNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);

            var fileDirectoryPath = Path.Combine(outputDirectory, "temp", fileNameWithoutExt);

            ClearFolder(fileDirectoryPath);

            var converter = GetKompasFilePngConverter(fileType);
            converter.ConvertToRaster(filePath, fileDirectoryPath);

            return fileDirectoryPath;
        }

        private static void ClearFolder(string fileDirectoryPath)
        {
            if (Directory.Exists(fileDirectoryPath))
            {
                var directoryInfo = new DirectoryInfo(fileDirectoryPath);

                foreach (var file in directoryInfo.GetFiles())
                {
                    file.Delete();
                }
            }
        }

        private BaseKompasPngConverter GetKompasFilePngConverter(KompasFileType fileType)
        {
            // ReSharper disable once SwitchExpressionHandlesSomeKnownEnumValuesWithExceptionInDefault
            return fileType switch
            {
                KompasFileType.Draw2D => new KompasDraw2DToPngConverter(_kompas),
                KompasFileType.Specification => new KompasSpecificationToPngConverter(_kompas),
                KompasFileType.Draw3D => new KompasDraw3DToPngConverter(_kompas),
                _ => throw new InvalidEnumArgumentException()
            };
        }

        private static KompasFileType GetKompasFileType(string filePath)
        {
            var extension = Path.GetExtension(filePath);

            return extension switch
            {
                ".cdw" => KompasFileType.Draw2D,
                ".spw" => KompasFileType.Specification,
                ".m3d" => KompasFileType.Draw3D,
                _ => KompasFileType.Unknown
            };
        }

        public void ConvertToDwg(string filePath, string outputDwgDirectory)
        {
            var fileType = GetKompasFileType(filePath);
            
            var converter = GetKompasFileDwgConverter(fileType);
            
            converter.ConvertToDwg(filePath, outputDwgDirectory);
        }
        
        private BaseKompasDwgConverter GetKompasFileDwgConverter(KompasFileType fileType)
        {
            // ReSharper disable once SwitchExpressionHandlesSomeKnownEnumValuesWithExceptionInDefault
            return fileType switch
            {
                KompasFileType.Draw2D => new KompasDraw2DToDwgConverter(_kompas),
                KompasFileType.Specification => new KompasSpecificationToDwgConverter(_kompas),
                _ => throw new InvalidEnumArgumentException()
            };
        }
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                _instance = null;
            }
            _kompas.Quit();
            Marshal.ReleaseComObject(_kompas);
        }

        ~KompasInstance()
        {
            Dispose(false);
        }

        public static bool IsNotNullInstance()
        {
            return _instance != null;
        }
    }

    public enum KompasFileType
    {
        Unknown = 0,
        Draw2D = 1,
        Specification = 2,
        Draw3D = 3,
        Assembly = 4
    }
}