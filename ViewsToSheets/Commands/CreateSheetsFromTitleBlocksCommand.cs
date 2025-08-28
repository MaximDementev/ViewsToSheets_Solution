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

                    int successSheetsCount = CreateSheetsFromTitleBlockGroups(doc, activeSheet, titleBlockGroups);

                    if (successSheetsCount > 0)
                    {
                        trans.Commit();
                        TaskDialog.Show(Messages.SUCCESS_TITLE,
                            string.Format(Messages.SHEETS_CREATED_SUCCESS, successSheetsCount));
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

                    bool isIntersect = vpOutline.Intersects(tbOutline, 0.0);
                    bool isVpInTb = IsOutlineCenterInside(vpOutline, tbOutline);

                    // Проверяем пересечение прямоугольников
                    if (isIntersect || isVpInTb)
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

        private bool IsOutlineCenterInside(Outline inner, Outline outer)
        {
            if (inner == null || outer == null)
                return false;

            // Центр по XY
            XYZ center = (inner.MinimumPoint + inner.MaximumPoint) / 2;

            return
                center.X >= outer.MinimumPoint.X &&
                center.X <= outer.MaximumPoint.X &&
                center.Y >= outer.MinimumPoint.Y &&
                center.Y <= outer.MaximumPoint.Y;
        }


        private Outline GetTitleBlockOutline(FamilyInstance titleBlock)
        {
            LocationPoint loc = titleBlock.Location as LocationPoint;
            if (loc == null) return null;

            XYZ origin = loc.Point; // точка вставки (нижний правый угол)

            double width = titleBlock.get_Parameter(BuiltInParameter.SHEET_WIDTH).AsDouble();
            double height = titleBlock.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble();

            // Нижний левый и верхний правый
            XYZ min = new XYZ(origin.X - width, origin.Y, 0);
            XYZ max = new XYZ(origin.X, origin.Y + height, 0);

            return new Outline(min, max);
        }


        /// <summary>
        /// Создает новые листы для каждой группы основных надписей и переносит виды.
        /// </summary>
        private int CreateSheetsFromTitleBlockGroups(Document doc, ViewSheet originalSheet,
            Dictionary<FamilyInstance, List<Viewport>> titleBlockGroups)
        {
            int successSheetsCount = 0;
            try
            {
                var groupsList = titleBlockGroups.ToList();

                // Оставляем первую группу на исходном листе, создаем новые листы для остальных
                for (int i = 1; i < groupsList.Count; i++)
                {
                    FamilyInstance originalTitleBlock = groupsList[i].Key;
                    List<Viewport> viewports = groupsList[i].Value;
                    if (viewports.Count == 0) continue;

                    // Создаем новый лист
                    ViewSheet newSheet = CreateNewSheet(doc, originalSheet, originalTitleBlock, viewports, i);
                    if (newSheet == null) continue;


                    // Переносим виды на новый лист с сохранением относительного положения
                    MoveViewportsToSheet(doc, titleBlockGroups, i, newSheet);


                    // Копируем параметры листа
                    ParameterCopyService.CopySheetParameters(originalSheet, newSheet);
                    // Копируем параметры основной надписи
                    var targetTitleBlock = GetTitleBlocksOnSheet(doc, newSheet).FirstOrDefault();
                    ParameterCopyService.CopyTitleBlockParameters(originalTitleBlock, targetTitleBlock);

                    successSheetsCount++;
                }

                // Удаляем лишние основные надписи с исходного листа
                for (int i = 1; i < groupsList.Count; i++)
                {
                    if(groupsList[i].Value.Count !=0)
                    doc.Delete(groupsList[i].Key.Id);
                }
                return successSheetsCount;
            }
            catch
            {
                return successSheetsCount;
            }
        }

        /// <summary>
        /// Создает новый лист на основе исходного.
        /// </summary>
        private ViewSheet CreateNewSheet(Document doc, ViewSheet originalSheet, FamilyInstance titleBlock, List<Viewport> viewports, int index)
        {
            try
            {
                var sheetTypeId = originalSheet.GetTypeId();
                var sheetType = doc.GetElement(sheetTypeId) as ElementType;

                if (sheetType == null) return null;

                // Создаем уникальный номер и имя листа
                string sheetName = GetViewsNamesString(doc, viewports);
                sheetName = GetUniqueViewName(doc, sheetName);
                string sheetNumber = $"{originalSheet.SheetNumber}-{index}";
                sheetNumber = GetUniqueSheetNumber(doc, sheetNumber);

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

        private string GetUniqueViewName(Document doc, string baseName)
        {
            var existingNames = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Select(v => v.Name)
                .ToHashSet();

            if (!existingNames.Contains(baseName))
                return baseName;

            int counter = 1;
            string newName;
            do
            {
                newName = $"{baseName} ({counter})";
                counter++;
            }
            while (existingNames.Contains(newName));

            return newName;
        }

        private string GetUniqueSheetNumber(Document doc, string baseNumber)
        {
            var existingNumbers = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Select(s => s.SheetNumber)
                .ToHashSet();

            if (!existingNumbers.Contains(baseNumber))
                return baseNumber;

            int counter = 1;
            string newNumber;
            do
            {
                newNumber = $"{baseNumber} ({counter})";
                counter++;
            }
            while (existingNumbers.Contains(newNumber));

            return newNumber;
        }

        private string GetViewsNamesString(Document doc, List<Viewport> viewports)
        {
            var names = viewports
                .Select(vp => doc.GetElement(vp.ViewId) as View)
                .Where(v => v != null)
                .Select(v => v.Name)
                .ToList();

            return string.Join(". ", names);
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
