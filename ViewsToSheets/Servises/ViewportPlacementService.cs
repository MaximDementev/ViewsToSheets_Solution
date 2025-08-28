using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagicEntry.Plugins.ViewsToSheets.Services
{
    /// <summary>
    /// Сервис для размещения viewports на листах.
    /// Предоставляет методы для размещения видов в различных конфигурациях.
    /// </summary>
    public class ViewportPlacementService
    {
        #region Public Methods

        /// <summary>
        /// Размещает виды в столбик справа от общего контура элементов на листе.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="views">Список видов для размещения</param>
        /// <param name="sheet">Целевой лист</param>
        /// <param name="spacingMm">Расстояние между видами в мм</param>
        /// <returns>True в случае успешного размещения</returns>
        public bool PlaceViewsInColumn(Document doc, List<View> views, ViewSheet sheet, double spacingMm)
        {
            if (doc == null || views == null || sheet == null || !views.Any())
                return false;

            try
            {
                var spacingFeet = spacingMm * Constants.MM_TO_FEET;

                // Получаем все элементы на листе
                var (viewports, titleBlocks) = GetSheetElements(doc, sheet);
                var combinedOutline = GetCombinedOutline(doc, viewports, titleBlocks);
                if (combinedOutline == null) return false;

                // Вычисляем позицию для размещения
                double rightEdgeX = combinedOutline.MaximumPoint.X;
                double currentY = combinedOutline.MaximumPoint.Y;
                double leftAlignmentX = rightEdgeX + Constants.SPACING_LARGE / Constants.FEET_TO_MM;

                return PlaceViewsAtPosition(doc, views, sheet, leftAlignmentX, currentY, spacingFeet);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Переносит viewports на новый лист с сохранением относительного положения.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="titleBlockGroups">Группы основных надписей и viewports</param>
        /// <param name="groupIndex">Индекс группы для переноса</param>
        /// <param name="targetSheet">Целевой лист</param>
        public void MoveViewportsToSheet(Document doc, Dictionary<FamilyInstance, List<Viewport>> titleBlockGroups,
            int groupIndex, ViewSheet targetSheet)
        {
            if (doc == null || titleBlockGroups == null || targetSheet == null ||
                groupIndex < 0 || groupIndex >= titleBlockGroups.Count) return;

            var pair = titleBlockGroups.ElementAt(groupIndex);
            FamilyInstance titleBlock = pair.Key;
            List<Viewport> viewports = pair.Value;

            if (!(titleBlock.Location is LocationPoint loc)) return;

            XYZ basePoint = loc.Point; // нижний правый угол рамки

            foreach (var viewport in viewports)
            {
                MoveViewportToSheet(doc, viewport, targetSheet, basePoint);
            }
        }

        #endregion

        #region Private Methods

        /// <summary>
        /// Получает элементы листа (viewports и основные надписи).
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="sheet">Лист</param>
        /// <returns>Кортеж с viewports и основными надписями</returns>
        private (List<Viewport> viewports, List<FamilyInstance> titleBlocks) GetSheetElements(Document doc, ViewSheet sheet)
        {
            var viewService = new ViewService();
            var titleBlockService = new TitleBlockService();

            var viewports = viewService.GetViewportsOnSheet(doc, sheet);
            var titleBlocks = titleBlockService.GetTitleBlocksOnSheet(doc, sheet);

            return (viewports, titleBlocks);
        }

        /// <summary>
        /// Получает общий габарит всех элементов на листе.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="viewports">Список viewports</param>
        /// <param name="titleBlocks">Список основных надписей</param>
        /// <returns>Общий габарит или null</returns>
        private Outline GetCombinedOutline(Document doc, List<Viewport> viewports, List<FamilyInstance> titleBlocks)
        {
            List<Outline> outlines = new List<Outline>();
            var titleBlockService = new TitleBlockService();

            // Добавляем габариты viewports
            outlines.AddRange(viewports.Select(vp => vp.GetBoxOutline()).Where(outline => outline != null));

            // Добавляем габариты основных надписей
            outlines.AddRange(titleBlocks.Select(tb => titleBlockService.GetTitleBlockOutline(tb)).Where(outline => outline != null));

            if (!outlines.Any()) return null;

            // Вычисляем общие границы
            double minX = outlines.Min(o => o.MinimumPoint.X);
            double minY = outlines.Min(o => o.MinimumPoint.Y);
            double maxX = outlines.Max(o => o.MaximumPoint.X);
            double maxY = outlines.Max(o => o.MaximumPoint.Y);

            return new Outline(new XYZ(minX, minY, 0), new XYZ(maxX, maxY, 0));
        }

        /// <summary>
        /// Размещает виды в указанной позиции.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="views">Список видов</param>
        /// <param name="sheet">Лист</param>
        /// <param name="startX">Начальная X координата</param>
        /// <param name="startY">Начальная Y координата</param>
        /// <param name="spacingFeet">Расстояние между видами в футах</param>
        /// <returns>True в случае успеха</returns>
        private bool PlaceViewsAtPosition(Document doc, List<View> views, ViewSheet sheet,
            double startX, double startY, double spacingFeet)
        {
            double currentY = startY;

            foreach (var view in views)
            {
                if (!Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                    continue;

                // Создаём viewport в временной позиции
                var tempPosition = new XYZ(0, currentY, 0);
                var viewport = Viewport.Create(doc, sheet.Id, view.Id, tempPosition);
                if (viewport == null) continue;

                // Вычисляем финальную позицию
                var outline = viewport.GetBoxOutline();
                double viewportHeight = outline.MaximumPoint.Y - outline.MinimumPoint.Y;
                double viewportLeft = outline.MinimumPoint.X;
                double offsetX = startX - viewportLeft;

                var finalPosition = new XYZ(tempPosition.X + offsetX, currentY, 0);
                viewport.SetBoxCenter(finalPosition);

                // Смещаемся для следующего вида
                currentY -= viewportHeight + spacingFeet;
            }

            return true;
        }

        /// <summary>
        /// Переносит отдельный viewport на новый лист.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="viewport">Viewport для переноса</param>
        /// <param name="targetSheet">Целевой лист</param>
        /// <param name="basePoint">Базовая точка для расчета относительного положения</param>
        private void MoveViewportToSheet(Document doc, Viewport viewport, ViewSheet targetSheet, XYZ basePoint)
        {
            View view = doc.GetElement(viewport.ViewId) as View;
            if (view == null) return;

            // Вычисляем относительное положение
            Outline outline = viewport.GetBoxOutline();
            XYZ max = outline.MaximumPoint;
            double width = outline.MaximumPoint.X - outline.MinimumPoint.X;
            double height = outline.MaximumPoint.Y - outline.MinimumPoint.Y;

            XYZ relativeOffset = max - basePoint;
            XYZ newPos = relativeOffset - new XYZ(width / 2, height / 2, 0);

            // Удаляем старый viewport
            doc.Delete(viewport.Id);

            // Создаем новый viewport
            if (Viewport.CanAddViewToSheet(doc, targetSheet.Id, view.Id))
            {
                Viewport newVp = Viewport.Create(doc, targetSheet.Id, view.Id, newPos);
                if (newVp == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Ошибка: не удалось создать viewport для вида {view.Name}");
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"Вид {view.Name} нельзя добавить на лист {targetSheet.Name}");
            }
        }

        #endregion
    }
}
