using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace MagicEntry.Plugins.ViewsToSheets.Services
{
    /// <summary>
    /// Сервис для работы с основными надписями в Revit.
    /// Предоставляет методы для получения, анализа и группировки основных надписей.
    /// </summary>
    public class TitleBlockService
    {
        #region Public Methods

        /// <summary>
        /// Получает все основные надписи на указанном листе.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="sheet">Лист для анализа</param>
        /// <returns>Список основных надписей на листе</returns>
        public List<FamilyInstance> GetTitleBlocksOnSheet(Document doc, ViewSheet sheet)
        {
            if (doc == null || sheet == null) return new List<FamilyInstance>();

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
        /// Группирует виды по основным надписям на основе пересечения габаритов.
        /// </summary>
        /// <param name="doc">Документ Revit</param>
        /// <param name="titleBlocks">Список основных надписей</param>
        /// <param name="viewports">Список viewports</param>
        /// <returns>Словарь групп: основная надпись -> список viewports</returns>
        public Dictionary<FamilyInstance, List<Viewport>> GroupViewportsByTitleBlocks(
            Document doc, List<FamilyInstance> titleBlocks, List<Viewport> viewports)
        {
            if (doc == null || titleBlocks == null || viewports == null)
                return new Dictionary<FamilyInstance, List<Viewport>>();

            var groups = new Dictionary<FamilyInstance, List<Viewport>>();

            // Инициализируем группы
            foreach (var titleBlock in titleBlocks)
                groups[titleBlock] = new List<Viewport>();

            // Перебираем виды
            foreach (var viewport in viewports)
            {
                Outline vpOutline = viewport.GetBoxOutline();
                FamilyInstance containingTitleBlock = FindContainingTitleBlock(titleBlocks, vpOutline);

                if (containingTitleBlock != null)
                    groups[containingTitleBlock].Add(viewport);
            }

            return groups;
        }

        /// <summary>
        /// Получает габариты основной надписи.
        /// </summary>
        /// <param name="titleBlock">Основная надпись</param>
        /// <returns>Габариты основной надписи</returns>
        public Outline GetTitleBlockOutline(FamilyInstance titleBlock)
        {
            if (titleBlock == null) return null;

            var loc = titleBlock.Location as LocationPoint;
            if (loc == null) return null;

            XYZ origin = loc.Point; // точка вставки (нижний правый угол)
            double rotation = loc.Rotation; // угол поворота

            double width = titleBlock.get_Parameter(BuiltInParameter.SHEET_WIDTH).AsDouble();
            double height = titleBlock.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble();

            if (Math.Abs(rotation) < 1e-6) // без поворота
            {
                XYZ min = new XYZ(origin.X - width, origin.Y, 0);
                XYZ max = new XYZ(origin.X, origin.Y + height, 0);
                return new Outline(min, max);
            }

            // С учетом поворота
            return GetRotatedTitleBlockOutline(origin, width, height, rotation);
        }


        #endregion

        #region Private Methods

        /// <summary>
        /// Находит основную надпись, содержащую указанный viewport.
        /// </summary>
        /// <param name="titleBlocks">Список основных надписей</param>
        /// <param name="vpOutline">Габариты viewport</param>
        /// <returns>Содержащая основная надпись или null</returns>
        private FamilyInstance FindContainingTitleBlock(List<FamilyInstance> titleBlocks, Outline vpOutline)
        {
            foreach (var titleBlock in titleBlocks)
            {
                Outline tbOutline = GetTitleBlockOutline(titleBlock);
                if (tbOutline == null) continue;

                bool isIntersect = vpOutline.Intersects(tbOutline, 0.0);
                bool isVpInTb = IsOutlineCenterInside(vpOutline, tbOutline);

                if (isIntersect || isVpInTb)
                    return titleBlock;
            }

            return null;
        }

        /// <summary>
        /// Проверяет, находится ли центр одного габарита внутри другого.
        /// </summary>
        /// <param name="inner">Внутренний габарит</param>
        /// <param name="outer">Внешний габарит</param>
        /// <returns>True, если центр inner находится внутри outer</returns>
        private bool IsOutlineCenterInside(Outline inner, Outline outer)
        {
            if (inner == null || outer == null) return false;

            XYZ center = (inner.MinimumPoint + inner.MaximumPoint) / 2;

            return center.X >= outer.MinimumPoint.X &&
                   center.X <= outer.MaximumPoint.X &&
                   center.Y >= outer.MinimumPoint.Y &&
                   center.Y <= outer.MaximumPoint.Y;
        }

        /// <summary>
        /// Получает габариты повернутой основной надписи.
        /// </summary>
        /// <param name="origin">Точка вставки</param>
        /// <param name="width">Ширина</param>
        /// <param name="height">Высота</param>
        /// <param name="rotation">Угол поворота</param>
        /// <returns>Габариты повернутой основной надписи</returns>
        private Outline GetRotatedTitleBlockOutline(XYZ origin, double width, double height, double rotation)
        {
            // Локальные точки (в системе координат рамки)
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
}
