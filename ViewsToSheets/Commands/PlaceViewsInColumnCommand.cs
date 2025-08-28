using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using MagicEntry.Core.Interfaces;
using MagicEntry.Core.Models;
using MagicEntry.Core.Services;
using MagicEntry.Plugins.ViewsToSheets.Services;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagicEntry.Plugins.ViewsToSheets
{
    /// <summary>
    /// Команда для размещения выбранных видов в столбик на одном листе.
    /// Позволяет выбрать виды из диспетчера проекта и разместить их на указанном листе с заданным расстоянием.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class PlaceViewsInColumnCommand : IExternalCommand, IPlugin
    {
        #region Fields

        private readonly ViewService _viewService;
        private readonly ViewportPlacementService _placementService;

        #endregion

        #region Constructor

        /// <summary>
        /// Инициализирует новый экземпляр команды размещения видов в столбик.
        /// </summary>
        public PlaceViewsInColumnCommand()
        {
            _viewService = new ViewService();
            _placementService = new ViewportPlacementService();
        }

        #endregion

        #region IPlugin Implementation

        /// <summary>
        /// Конфигурационная информация о плагине, устанавливается системой.
        /// </summary>
        public PluginInfo Info { get; set; }

        /// <summary>
        /// Указывает, активен ли плагин.
        /// </summary>
        public bool IsEnabled { get; set; }

        /// <summary>
        /// Выполняет внутреннюю инициализацию плагина при загрузке системы.
        /// </summary>
        /// <returns>True в случае успешной инициализации</returns>
        public bool Initialize()
        {
            try
            {
                var pathService = ServiceProvider.GetService<IPathService>();
                var initService = ServiceProvider.GetService<IPluginInitializationService>();

                if (pathService == null || initService == null)
                    return false;

                var pluginName = Info?.Name ?? Constants.PLUGIN_NAME;
                return initService.InitializePlugin(pluginName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Освобождает ресурсы, используемые плагином.
        /// </summary>
        public void Shutdown()
        {
            // Логика завершения работы плагина
        }

        #endregion

        #region IExternalCommand Implementation

        /// <summary>
        /// Точка входа для выполнения команды размещения видов в столбик.
        /// </summary>
        /// <param name="commandData">Данные команды</param>
        /// <param name="message">Сообщение об ошибке</param>
        /// <param name="elements">Набор элементов</param>
        /// <returns>Результат выполнения команды</returns>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiDoc = commandData.Application.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Получаем выбранные виды
                List<View> selectedViews = _viewService.GetSelectedViews(uiDoc);
                if (selectedViews == null || !selectedViews.Any())
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.NO_VIEWS_SELECTED);
                    return Result.Cancelled;
                }

                // Выбираем целевой лист
                ViewSheet targetSheet = GetTargetSheet(uiDoc, doc);
                if (targetSheet == null)
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, "Нужно выбрать лист или открыть его.");
                    return Result.Cancelled;
                }

                // Выполняем размещение видов
                return ExecuteViewPlacement(doc, selectedViews, targetSheet);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show(Messages.ERROR_TITLE, string.Format(Messages.EXECUTION_ERROR, ex.Message));
                return Result.Failed;
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Получает целевой лист для размещения видов.
        /// </summary>
        /// <param name="uiDoc">UI документ</param>
        /// <param name="doc">Документ</param>
        /// <returns>Целевой лист или null</returns>
        private ViewSheet GetTargetSheet(UIDocument uiDoc, Document doc)
        {
            // Пытаемся выбрать лист
            Reference pickedRef = null;
            try
            {
                pickedRef = uiDoc.Selection.PickObject(
                    ObjectType.Element,
                    new SheetSelectionFilter(),
                    "Выберите лист"
                );
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // Если отмена - используем активный вид
            }

            if (pickedRef != null)
            {
                return doc.GetElement(pickedRef) as ViewSheet;
            }
            else if (uiDoc.ActiveView is ViewSheet activeSheet)
            {
                return activeSheet;
            }

            return null;
        }

        /// <summary>
        /// Выполняет размещение видов на листе.
        /// </summary>
        /// <param name="doc">Документ</param>
        /// <param name="selectedViews">Выбранные виды</param>
        /// <param name="targetSheet">Целевой лист</param>
        /// <returns>Результат выполнения</returns>
        private Result ExecuteViewPlacement(Document doc, List<View> selectedViews, ViewSheet targetSheet)
        {
            double spacing = Constants.SPACING_LARGE;

            using (Transaction trans = new Transaction(doc, Messages.TRANSACTION_NAME))
            {
                trans.Start();

                bool success = _placementService.PlaceViewsInColumn(doc, selectedViews, targetSheet, spacing);

                if (success)
                {
                    trans.Commit();
                    TaskDialog.Show(Messages.SUCCESS_TITLE,
                        string.Format(Messages.VIEWS_PLACED_SUCCESS, selectedViews.Count, targetSheet.Name));
                    return Result.Succeeded;
                }
                else
                {
                    trans.RollBack();
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.PLACEMENT_FAILED);
                    return Result.Failed;
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// Фильтр для выбора только листов.
    /// </summary>
    public class SheetSelectionFilter : ISelectionFilter
    {
        #region ISelectionFilter Implementation

        /// <summary>
        /// Определяет, можно ли выбрать элемент.
        /// </summary>
        /// <param name="elem">Проверяемый элемент</param>
        /// <returns>True, если элемент является листом</returns>
        public bool AllowElement(Element elem)
        {
            return elem is ViewSheet;
        }

        /// <summary>
        /// Определяет, можно ли выбрать ссылку.
        /// </summary>
        /// <param name="reference">Ссылка</param>
        /// <param name="position">Позиция</param>
        /// <returns>False - ссылки не разрешены</returns>
        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }

        #endregion
    }
}
