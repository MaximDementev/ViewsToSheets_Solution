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
                List<View> selectedViews = GetSelectedViews(uiDoc);
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
                double spacing = Constants.SPACING_LARGE;

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
        /// Размещает виды в столбик справа от общего контура элементов на листе, 
        /// с выравниванием по левой грани.
        /// </summary>
        private bool PlaceViewsInColumn(Document doc, List<View> views, ViewSheet sheet, double spacingMm)
        {
            try
            {
                var spacingFeet = spacingMm * Constants.MM_TO_FEET;

                // Получаем все элементы на листе
                var (viewports, titleBlocks) = GetSheetElements(doc, sheet);
                var combinedOutline = GetCombinedOutline(doc, viewports, titleBlocks);
                if (combinedOutline == null)
                    return false;

                // Правая граница существующих элементов
                double rightEdgeX = combinedOutline.MaximumPoint.X;

                // Начальная позиция для первого вида
                double currentY = combinedOutline.MaximumPoint.Y; // сверху вниз
                double leftAlignmentX = rightEdgeX + Constants.SPACING_LARGE / Constants.FEET_TO_MM; // смещение вправо

                foreach (var view in views)
                {
                    if (!Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                        continue;

                    // Создаём в нуле, потом передвигаем
                    var tempPosition = new XYZ(0, currentY, 0);
                    var viewport = Viewport.Create(doc, sheet.Id, view.Id, tempPosition);
                    if (viewport == null)
                        continue;

                    // Размер вьюпорта
                    var outline = viewport.GetBoxOutline();
                    double viewportHeight = outline.MaximumPoint.Y - outline.MinimumPoint.Y;
                    double viewportLeft = outline.MinimumPoint.X;

                    // Выравнивание по левой грани
                    double offsetX = leftAlignmentX - viewportLeft;

                    // Итоговая позиция центра
                    var finalPosition = new XYZ(tempPosition.X + offsetX, currentY, 0);
                    viewport.SetBoxCenter(finalPosition);

                    // Смещаемся вниз на высоту + зазор
                    currentY -= viewportHeight + spacingFeet;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private (List<Viewport> viewports, List<FamilyInstance> titleBlocks) GetSheetElements(Document doc, ViewSheet sheet)
        {
            var viewports = new FilteredElementCollector(doc, sheet.Id)
                .OfClass(typeof(Viewport))
                .Cast<Viewport>()
                .ToList();

            var titleBlocks = new FilteredElementCollector(doc, sheet.Id)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Category?.Id.IntegerValue == (int)BuiltInCategory.OST_TitleBlocks)
                .ToList();

            return (viewports, titleBlocks);
        }

        private Outline GetCombinedOutline(Document doc, List<Viewport> viewports, List<FamilyInstance> titleBlocks)
        {
            List<Outline> outlines = new List<Outline>();

            // Вьюпорты
            foreach (var vp in viewports)
            {
                Outline vpOutline = vp.GetBoxOutline();
                if (vpOutline != null)
                    outlines.Add(vpOutline);
            }

            // Основные надписи
            foreach (var tb in titleBlocks)
            {
                Outline tbOutline = GetTitleBlockOutline(tb);
                if (tbOutline != null)
                    outlines.Add(tbOutline);
            }

            if (!outlines.Any())
                return null;

            // Общие min/max по всем
            double minX = outlines.Min(o => o.MinimumPoint.X);
            double minY = outlines.Min(o => o.MinimumPoint.Y);
            double maxX = outlines.Max(o => o.MaximumPoint.X);
            double maxY = outlines.Max(o => o.MaximumPoint.Y);

            return new Outline(new XYZ(minX, minY, 0), new XYZ(maxX, maxY, 0));
        }

        private Outline GetTitleBlockOutline(FamilyInstance titleBlock)
        {
            LocationPoint loc = titleBlock.Location as LocationPoint;
            if (loc == null) return null;

            XYZ origin = loc.Point;              // точка вставки (внизу справа)
            double rotation = loc.Rotation;      // угол поворота

            double width = titleBlock.get_Parameter(BuiltInParameter.SHEET_WIDTH).AsDouble();
            double height = titleBlock.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble();

            // Локальные точки (в системе координат рамки)
            // Вставка = (0,0) нижний правый угол
            List<XYZ> localCorners = new List<XYZ>
                {
                    new XYZ(0, 0, 0),           // нижний правый
                    new XYZ(-width, 0, 0),      // нижний левый
                    new XYZ(-width, height, 0), // верхний левый
                    new XYZ(0, height, 0)       // верхний правый
                };

            // Матрица поворота вокруг Z
            Transform rot = Transform.CreateRotationAtPoint(XYZ.BasisZ, rotation, origin);

            // Применяем трансформацию к углам
            List<XYZ> worldCorners = localCorners.Select(p => rot.OfPoint(p + origin)).ToList();

            // Строим Outline по углам
            XYZ min = new XYZ(
                worldCorners.Min(p => p.X),
                worldCorners.Min(p => p.Y),
                0);
            XYZ max = new XYZ(
                worldCorners.Max(p => p.X),
                worldCorners.Max(p => p.Y),
                0);

            return new Outline(min, max);
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
