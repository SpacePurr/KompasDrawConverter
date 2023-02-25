using System;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using KompasDrawConverter.Providers;
using Spectre.Console;

namespace KompasDrawConverter
{
    [SuppressMessage("ReSharper", "StringLiteralTypo")]
    public static class Program
    {
        private static ConsoleEventDelegate _handler;
        
        private delegate bool ConsoleEventDelegate(int eventType);
        
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleCtrlHandler(ConsoleEventDelegate callback, bool add);
        
        private static bool ConsoleEventCallback(int eventType)
        {
            if (eventType == 2)
            {
                OnProcessExit();
            }
            return false;
        }

        private static void OnProcessExit()
        {
            if (KompasInstance.IsNotNullInstance())
            {
                KompasInstance.GetInstance().Dispose();
            }
        }

        public static void Main(string[] args)
        {
            _handler = ConsoleEventCallback;
            SetConsoleCtrlHandler(_handler, true);
            
            while (true)
            {
                try
                {
                    var converterMode = GetConverterMode();

                    if (converterMode == null)
                    {
                        break;
                    }

                    var inputDirectory = GetInputDirectory();

                    var options = new Options
                    {
                        InputDirectory = inputDirectory,
                        OutputDirectory = Path.Combine(inputDirectory),
                        ConverterMode = converterMode.Value
                    };

                    var converter = GetKompasFileConverter(options);
                    converter.Convert();
                }
                catch (Exception ex)
                {
                    AnsiConsole.WriteException(ex);
                }
            }
            
            AnsiConsole.Write(
                new FigletText("Kompas Draw Converter")
                    .LeftJustified()
                    .Color(Color.Blue3));

            Task.Delay(2000).Wait();
        }
        
        private static KompasFileConverter GetKompasFileConverter(Options options)
        {
            return options.ConverterMode switch
            {
                ConverterMode.ToPdf => new KompasToPngConverterProvider(options),
                ConverterMode.ToDwg => new KompasToDwgConverterProvider(options),
                _ => throw new InvalidEnumArgumentException(options.ConverterMode.ToString())
            };
        }

        private static ConverterMode? GetConverterMode()
        {
            var promt = new SelectionPrompt<string>()
                .Title("[green]Выберите режим работы программы[/]")
                .PageSize(10)
                .MoreChoicesText("[grey](Нажмите вверх или вниз для движения по списку)[/]")
                .AddChoices(Constants.LocalizationConverterMode.Keys.ToArray())
                .AddChoices("Выход");
            
            var converterModeLcz = AnsiConsole.Prompt(promt);

            if (Constants.LocalizationConverterMode.TryGetValue(converterModeLcz, out var converterMode))
            {
                return converterMode;
            }

            return null;
        }

        private static string GetInputDirectory()
        {
            var promt = new TextPrompt<string>("Введите путь до папки с чертежам")
                .PromptStyle("green")
                .Validate(inputDirectory =>
                    Directory.Exists(inputDirectory)
                        ? ValidationResult.Success()
                        : ValidationResult.Error($"[red]Путь {inputDirectory} не найден[/]"));
                
            return AnsiConsole.Prompt(promt);
        }
    }

    public enum ConverterMode
    {
        ToPdf = 1,
        ToDwg = 2
    }
}