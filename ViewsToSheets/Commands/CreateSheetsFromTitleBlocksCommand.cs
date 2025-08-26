using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using MagicEntry.Core.Interfaces;
using MagicEntry.Core.Models;
using MagicEntry.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using ViewsToSheets.Servises;

namespace MagicEntry.Plugins.ViewsToSheets
{
    /// <summary>
    /// Команда для создания отдельных листов для каждой основной надписи.
    /// Анализирует активный лист, находит основные надписи и виды внутри них,
    /// создает новые листы и переносит виды с сохранением относительного положения.
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateSheetsFromTitleBlocksCommand : IExternalCommand, IPlugin
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
        /// Точка входа для выполнения команды создания листов из основных надписей.
        /// </summary>
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                var uiDoc = commandData.Application.ActiveUIDocument;
                var doc = uiDoc.Document;

                // Проверяем, что активный вид - это лист
                if (!(uiDoc.ActiveView is ViewSheet activeSheet))
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.ACTIVE_VIEW_NOT_SHEET);
                    return Result.Cancelled;
                }

                // Получаем основные надписи на листе
                var titleBlocks = GetTitleBlocksOnSheet(doc, activeSheet);
                if (titleBlocks == null || !titleBlocks.Any())
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.NO_TITLE_BLOCKS_FOUND);
                    return Result.Cancelled;
                }

                // Получаем виды на листе
                var viewports = GetViewportsOnSheet(doc, activeSheet);
                if (viewports == null || !viewports.Any())
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.NO_VIEWS_ON_SHEET);
                    return Result.Cancelled;
                }

                // Группируем виды по основным надписям
                var titleBlockGroups = GroupViewportsByTitleBlocks(doc, titleBlocks, viewports);
                if (titleBlockGroups.Count <= 1)
                {
                    TaskDialog.Show(Messages.ERROR_TITLE, Messages.INSUFFICIENT_TITLE_BLOCKS);
                    return Result.Cancelled;
                }

                // Создаем новые листы
                using (Transaction trans = new Transaction(doc, Messages.TRANSACTION_NAME_COMMAND3))
                {
                    trans.Start();

                    bool success = CreateSheetsFromTitleBlockGroups(doc, activeSheet, titleBlockGroups);

                    if (success)
                    {
                        trans.Commit();
                        TaskDialog.Show(Messages.SUCCESS_TITLE,
                            string.Format(Messages.SHEETS_CREATED_SUCCESS, titleBlockGroups.Count - 1));
                        return Result.Succeeded;
                    }
                    else
                    {
                        trans.RollBack();
                        TaskDialog.Show(Messages.ERROR_TITLE, Messages.SHEET_CREATION_FAILED);
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
        /// Получает все основные надписи на указанном листе.
        /// </summary>
        private List<FamilyInstance> GetTitleBlocksOnSheet(Document doc, ViewSheet sheet)
        {
            var titleBlocks = new List<FamilyInstance>();

            var collector = new FilteredElementCollector(doc, sheet.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(typeof(FamilyInstance));

            foreach (FamilyInstance titleBlock in collector)
            {
                titleBlocks.Add(titleBlock);
            }

            return titleBlocks;
        }

        /// <summary>
        /// Получает все виды (viewports) на указанном листе.
        /// </summary>
        private List<Viewport> GetViewportsOnSheet(Document doc, ViewSheet sheet)
        {
            var viewports = new List<Viewport>();

            var viewportIds = sheet.GetAllViewports();
            foreach (var viewportId in viewportIds)
            {
                var viewport = doc.GetElement(viewportId) as Viewport;
                if (viewport != null)
                {
                    viewports.Add(viewport);
                }
            }

            return viewports;
        }

        /// <summary>
        /// Группирует виды по основным надписям на основе пересечения габаритов.
        /// Виды, которые не попадают в рамку, игнорируются.
        /// </summary>
        private Dictionary<FamilyInstance, List<Viewport>> GroupViewportsByTitleBlocks(
            Document doc, List<FamilyInstance> titleBlocks, List<Viewport> viewports)
        {
            var groups = new Dictionary<FamilyInstance, List<Viewport>>();

            // Инициализируем группы
            foreach (var titleBlock in titleBlocks)
                groups[titleBlock] = new List<Viewport>();

            // Перебираем виды
            foreach (var viewport in viewports)
            {
                Outline vpOutline = viewport.GetBoxOutline();

                FamilyInstance containingTitleBlock = null;
                foreach (var titleBlock in titleBlocks)
                {
                    Outline tbOutline = GetTitleBlockOutline(titleBlock);

                    // Проверяем пересечение прямоугольников
                    if (vpOutline.Intersects(tbOutline, 0.0))
                    {
                        containingTitleBlock = titleBlock;
                        break;
                    }
                }

                // Если не нашли рамку — пропускаем
                if (containingTitleBlock != null)
                    groups[containingTitleBlock].Add(viewport);
            }

            return groups;
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

        /// <summary>
        /// Создает новые листы для каждой группы основных надписей и переносит виды.
        /// </summary>
        private bool CreateSheetsFromTitleBlockGroups(Document doc, ViewSheet originalSheet,
            Dictionary<FamilyInstance, List<Viewport>> titleBlockGroups)
        {
            try
            {
                var groupsList = titleBlockGroups.ToList();

                // Оставляем первую группу на исходном листе, создаем новые листы для остальных
                for (int i = 1; i < groupsList.Count; i++)
                {
                    var originalTitleBlock = groupsList[i].Key;
                    var viewports = groupsList[i].Value;

                    // Создаем новый лист
                    ViewSheet newSheet = CreateNewSheet(doc, originalSheet, originalTitleBlock, i);
                    if (newSheet == null) continue;


                    // Переносим виды на новый лист с сохранением относительного положения
                    MoveViewportsToSheet(doc, titleBlockGroups, i, newSheet);


                    // Копируем параметры листа
                    ParameterCopyService.CopySheetParameters(originalSheet, newSheet);
                    // Копируем параметры основной надписи
                    var targetTitleBlock = GetTitleBlocksOnSheet(doc, newSheet).FirstOrDefault();
                    ParameterCopyService.CopyTitleBlockParameters(originalTitleBlock, targetTitleBlock);
                }

                // Удаляем лишние основные надписи с исходного листа
                for (int i = 1; i < groupsList.Count; i++)
                {
                    doc.Delete(groupsList[i].Key.Id);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Создает новый лист на основе исходного.
        /// </summary>
        private ViewSheet CreateNewSheet(Document doc, ViewSheet originalSheet, FamilyInstance titleBlock, int index)
        {
            try
            {
                var sheetTypeId = originalSheet.GetTypeId();
                var sheetType = doc.GetElement(sheetTypeId) as ElementType;

                if (sheetType == null) return null;

                // Создаем уникальный номер и имя листа
                var sheetNumber = $"{originalSheet.SheetNumber}-{index}";
                var sheetName = $"{originalSheet.Name} - Часть {index}";

                // Берем ID типа основной надписи из оригинального листа
                ElementId titleBlockTypeId = titleBlock.GetTypeId();
                if (titleBlockTypeId == ElementId.InvalidElementId)
                {
                    // Берем случайный тип основной надписи
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    titleBlockTypeId = collector
                        .OfCategory(BuiltInCategory.OST_TitleBlocks)
                        .OfClass(typeof(FamilySymbol))
                        .FirstElementId();
                }

                ViewSheet newSheet = ViewSheet.Create(doc, titleBlockTypeId);
                if (newSheet == null) return null;

                newSheet.SheetNumber = sheetNumber;
                newSheet.Name = sheetName;

                return newSheet;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Ошибка создания листа: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Переносит виды из i-й группы (основная надпись + её виды)
        /// на новый лист с сохранением относительного положения
        /// относительно LocationPoint рамки.
        /// </summary>
        private void MoveViewportsToSheet(
            Document doc,
            Dictionary<FamilyInstance, List<Viewport>> titleBlockGroups,
            int i,
            ViewSheet targetSheet)
        {
            if (i < 0 || i >= titleBlockGroups.Count) return;

            // Берём i-ю пару (рамка + виды)
            var pair = titleBlockGroups.ElementAt(i);
            FamilyInstance titleBlock = pair.Key;
            List<Viewport> viewports = pair.Value;

            LocationPoint loc = titleBlock.Location as LocationPoint;
            if (loc == null) return;

            XYZ basePoint = loc.Point; // нижний правый угол рамки

            foreach (var viewport in viewports)
            {
                View view = doc.GetElement(viewport.ViewId) as View;
                if (view == null) continue;

                Outline outline = viewport.GetBoxOutline();
                XYZ min = outline.MinimumPoint;
                XYZ max = outline.MaximumPoint;

                // Размеры рамки вида
                double width = max.X - min.X;
                double height = max.Y - min.Y;

                // Относительное смещение (берем верхний правый угол)
                XYZ currentPos = max;
                XYZ relativeOffset = currentPos - basePoint;

                // Удаляем старый viewport
                doc.Delete(viewport.Id);

                // Новая позиция = базовая точка + смещение - ширина (влево) - высота (вниз)
                XYZ newPos = relativeOffset - new XYZ(width/2, height/2, 0);

                if (Viewport.CanAddViewToSheet(doc, targetSheet.Id, view.Id))
                {
                    Viewport newVp = Viewport.Create(doc, targetSheet.Id, view.Id, newPos);
                    if (newVp == null)
                        System.Diagnostics.Debug.WriteLine($"Ошибка: не удалось создать viewport для вида {view.Name}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Вид {view.Name} нельзя добавить на лист {targetSheet.Name}");
                }
            }

        }



        #endregion
    }
}
