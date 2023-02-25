using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using Spectre.Console;

namespace KompasDrawConverter.Providers
{
    [SuppressMessage("ReSharper", "ConvertToUsingDeclaration")]
    public abstract class KompasFileConverter
    {
        private Options Options { get; }

        protected KompasFileConverter(Options options)
        {
            Options = options;
        }
        
        public void Convert()
        {
            var files = GetFiles();
            
            if (files.Length == 0)
            {
                AnsiConsole.MarkupLine($"В каталоге отсутствуют файлы в форматах [green]{string.Join(";", GetAllowedFilesExtensions())}[/]");
                return;
            }
            
            AnsiConsole.MarkupLine($"Количество файлов: [green]{files.Length}[/]");

            CreateDirectories();
            
            Convert(files);
        }

        protected abstract void CreateDirectories();
        

        protected abstract string[] GetAllowedFilesExtensions();

        private string[] GetFiles()
        {
            var allowedExtensions = GetAllowedFilesExtensions();
            
            var files = Directory
                .EnumerateFiles(Options.InputDirectory, "*.*")
                .Where(file => allowedExtensions.Contains(Path.GetExtension(file).TrimStart('.').ToLowerInvariant()))
                .ToArray();

            return files;
        }

        private void Convert(IReadOnlyCollection<string> files)
        {
            using (var kompasInstance = KompasInstance.GetInstance())
            {
                AnsiConsole.Progress()
                .Columns(
                    new TaskDescriptionColumn(),
                    new ProgressBarColumn(),        
                    new PercentageColumn(),
                    new RemainingTimeColumn(),
                    new SpinnerColumn())
                .Start(ctx =>
                {
                    var converterMode = Constants.LocalizationConverterMode
                        .FirstOrDefault(x => x.Value == Options.ConverterMode)
                        .Key;
                    
                    var task = ctx.AddTask(converterMode, new ProgressTaskSettings
                    {
                        AutoStart = false
                    });
                    
                    task.StartTask();
                    
                    // ReSharper disable once AccessToDisposedClosure
                    ProcessFiles(kompasInstance, files, task);
                });
            }
        }
        
        private void ProcessFiles(KompasInstance kompasInstance, IReadOnlyCollection<string> files, ProgressTask progressTask)
        {
            var increment = 100d / files.Count;
            
            foreach (var filePath in files)
            {
                var fileName = Path.GetFileName(filePath);
                    
                AnsiConsole.MarkupLine($"Обработка файла: [green]{fileName}[/]");

                try
                {
                    ConvertFile(kompasInstance, filePath);
                }
                catch (Exception ex)
                {
                    AnsiConsole.WriteException(ex);
                }
                finally
                {
                    progressTask.Increment(increment);
                }
            }
        }

        protected abstract void ConvertFile(KompasInstance kompasInstance, string filePath);
    }
}