using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using MagicEntry.Core.Interfaces;
using MagicEntry.Core.Models;
using MagicEntry.Core.Services;
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
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiDoc = commandData.Application.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Получаем выбранные виды
                var selectedViews = GetSelectedViews(uiDoc);
                if (selectedViews == null || !selectedViews.Any())
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.NO_VIEWS_SELECTED);
                    return Result.Cancelled;
                }

                // Выбираем целевой лист
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
                    // если Esc или отмена – используем активный вид
                }

                ViewSheet targetSheet = null;

                if (pickedRef != null)
                {
                    targetSheet = doc.GetElement(pickedRef) as ViewSheet;
                }
                else if (uiDoc.ActiveView is ViewSheet activeSheet)
                {
                    targetSheet = activeSheet;
                }

                if (targetSheet == null)
                {
                    TaskDialog.Show("Ошибка", "Нужно выбрать лист или открыть его.");
                    return Result.Cancelled;
                }

                // Запрашиваем расстояние между видами
                double spacing = GetSpacingFromUser();
                if (spacing <= 0)
                {
                    return Result.Cancelled;
                }

                // Размещаем виды на листе
                using (Transaction trans = new Transaction(doc, Messages.TRANSACTION_NAME))
                {
                    trans.Start();

                    bool success = PlaceViewsInColumn(doc, selectedViews, targetSheet, spacing);

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
        /// Получает выбранные в диспетчере проекта виды.
        /// </summary>
        private List<View> GetSelectedViews(UIDocument uiDoc)
        {
            var selection = uiDoc.Selection;
            var selectedIds = selection.GetElementIds();

            var views = new List<View>();
            foreach (var id in selectedIds)
            {
                var element = uiDoc.Document.GetElement(id);
                if (element is View view && CanBePlacedOnSheet(view))
                {
                    views.Add(view);
                }
            }

            return views;
        }

        /// <summary>
        /// Проверяет, может ли вид быть размещен на листе.
        /// </summary>
        private bool CanBePlacedOnSheet(View view)
        {
            return view.CanBePrinted &&
                   view.ViewType != ViewType.DrawingSheet &&
                   view.ViewType != ViewType.ProjectBrowser &&
                   view.ViewType != ViewType.SystemBrowser &&
                   !IsViewAlreadyOnSheet(view);
        }

        /// <summary>
        /// Проверяет, не размещен ли уже вид на каком-либо листе.
        /// </summary>
        private bool IsViewAlreadyOnSheet(View view)
        {
            try
            {
                var parameter = view.get_Parameter(BuiltInParameter.VIEWPORT_SHEET_NUMBER);
                return parameter != null && !string.IsNullOrEmpty(parameter.AsString());
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Запрашивает у пользователя расстояние между видами.
        /// </summary>
        private double GetSpacingFromUser()
        {
            var dialog = new TaskDialog(Messages.SPACING_DIALOG_TITLE);
            dialog.MainInstruction = Messages.SPACING_INSTRUCTION;
            dialog.MainContent = Messages.SPACING_DIALOG_CONTENT;

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                string.Format(Messages.SPACING_OPTION_SMALL, Constants.SPACING_SMALL));
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                string.Format(Messages.SPACING_OPTION_MEDIUM, Constants.SPACING_MEDIUM));
            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                string.Format(Messages.SPACING_OPTION_LARGE, Constants.SPACING_LARGE));

            dialog.CommonButtons = TaskDialogCommonButtons.Cancel;
            dialog.DefaultButton = TaskDialogResult.CommandLink2;

            var result = dialog.Show();

            switch (result)
            {
                case TaskDialogResult.CommandLink1:
                    return Constants.SPACING_SMALL;
                case TaskDialogResult.CommandLink2:
                    return Constants.SPACING_MEDIUM;
                case TaskDialogResult.CommandLink3:
                    return Constants.SPACING_LARGE;
                default:
                    return -1;
            }
        }

        /// <summary>
        /// Размещает виды в столбик на указанном листе с выравниванием по левой грани.
        /// </summary>
        private bool PlaceViewsInColumn(Document doc, List<View> views, ViewSheet sheet, double spacingMm)
        {
            try
            {
                var spacingFeet = spacingMm * Constants.MM_TO_FEET;
                var startPosition = new XYZ(Constants.START_POSITION_X, Constants.START_POSITION_Y, 0);
                var currentY = startPosition.Y;
                var leftAlignmentX = startPosition.X; // фиксированная X координата для выравнивания по левой грани

                foreach (var view in views)
                {
                    if (!Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                        continue;

                    var tempPosition = new XYZ(0, currentY, 0);
                    var viewport = Viewport.Create(doc, sheet.Id, view.Id, tempPosition);
                    if (viewport == null)
                        continue;

                    var outline = viewport.GetBoxOutline();
                    double viewportHeight = outline.MaximumPoint.Y - outline.MinimumPoint.Y;
                    double viewportLeft = outline.MinimumPoint.X;

                    double offsetX = leftAlignmentX - viewportLeft;
                    var finalPosition = new XYZ(tempPosition.X + offsetX, currentY, 0);

                    viewport.SetBoxCenter(finalPosition);

                    currentY = currentY - viewportHeight - spacingFeet;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }

    /// <summary>
    /// Фильтр для выбора только листов.
    /// </summary>
    public class SheetSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            return elem is ViewSheet;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }

}
