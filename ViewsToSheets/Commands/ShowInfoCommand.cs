using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ViewsToSheets.UI;

namespace ViewsToSheets.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ShowInfoCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                string assemblyLocation = Assembly.GetExecutingAssembly().Location;
                string assemblyDirectory = Path.GetDirectoryName(assemblyLocation);
                string infoFilePath = Path.Combine(assemblyDirectory, "ShowInfoCommand.txt");

                if (!File.Exists(infoFilePath))
                {
                    TaskDialog.Show("Ошибка", $"Файл справки 'ShowInfoCommand.txt' не найден в папке плагина:\n{assemblyDirectory}");
                    return Result.Failed;
                }

                string fileContent = File.ReadAllText(infoFilePath);

                // Парсинг текста и ссылок
                string mainText = ParseSection(fileContent, "##Текст", "##Ссылки");
                string linksRaw = ParseSection(fileContent, "##Ссылки");

                List<string> links = linksRaw.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                             .Select(l => l.Trim())
                                             .Where(l => !string.IsNullOrEmpty(l))
                                             .ToList();

                // Отображение окна
                using (var infoForm = new InfoWindowForm(mainText, links))
                {
                    infoForm.ShowDialog();
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private string ParseSection(string content, string startMarker, string endMarker = null)
        {
            int startIndex = content.IndexOf(startMarker);
            if (startIndex == -1) return string.Empty;

            startIndex += startMarker.Length;

            int endIndex = content.Length;
            if (endMarker != null)
            {
                endIndex = content.IndexOf(endMarker, startIndex);
                if (endIndex == -1) endIndex = content.Length;
            }

            return content.Substring(startIndex, endIndex - startIndex).Trim();
        }
    }
}
