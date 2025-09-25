using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System;
using System.Collections.Generic;

namespace TopoBuilder
{
    public class AlignmentCommands
    {
        [CommandMethod("ALTXT")]
        public void AlignTextToEntityCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Выбор образца (круг, дуга или блок)
                    Entity sampleEntity = SelectEntity(ed, tr);
                    Entity sampleText = SelectText(ed, tr);

                    // Поиск всех соответствующих объектов
                    var entities = FindMatchingEntities(tr, db, sampleEntity);
                    var texts = FindMatchingTexts(tr, db, sampleText);

                    int movedCount = 0;
                    var movedTexts = new HashSet<ObjectId>();
                    double tolerance = 0.001; // 1mm tolerance

                    // Вычисляем максимальный размер текста один раз
                    double maxTextDimension = CalculateMaxTextDimension(texts);

                    foreach (Entity text in texts)
                    {
                        if (movedTexts.Contains(text.ObjectId)) continue;

                        Entity closestEntity = null;
                        double minDistance = double.MaxValue;

                        // Поиск ближайшего объекта к тексту
                        foreach (Entity entity in entities)
                        {
                            Point3d center = GetEntityCenter(entity);
                            double distance = GetMinDistanceBetween(text, center);
                            if (distance < minDistance)
                            {
                                minDistance = distance;
                                closestEntity = entity;
                            }
                        }

                        if (closestEntity != null)
                        {
                            double searchRadius = CalculateSearchRadius(closestEntity, maxTextDimension);
                            if (minDistance <= searchRadius)
                            {
                                Point3d currentPosition = GetTextPosition(text);
                                Point3d targetCenter = GetEntityCenter(closestEntity);

                                // Проверка, не находится ли уже текст в целевой позиции
                                if (currentPosition.DistanceTo(targetCenter) > tolerance)
                                {
                                    MoveTextToPoint(text, targetCenter);
                                    movedCount++;
                                    movedTexts.Add(text.ObjectId);
                                }
                            }
                        }
                    }

                    tr.Commit();
                    ed.WriteMessage($"\nMoved {movedCount} text objects.");
                }
                catch (System.Exception ex)
                {
                    tr.Abort();
                    ed.WriteMessage($"\nError: {ex.Message}");
                }
            }
        }

        private Entity SelectEntity(Editor ed, Transaction tr)
        {
            var opts = new PromptEntityOptions("\nSelect sample circle, arc or block: ")
            {
                AllowNone = false,
                AllowObjectOnLockedLayer = true
            };
            opts.SetRejectMessage("Selected object is not a circle, arc or block");
            opts.AddAllowedClass(typeof(Circle), true);
            opts.AddAllowedClass(typeof(Arc), true);
            opts.AddAllowedClass(typeof(BlockReference), true);

            PromptEntityResult res = ed.GetEntity(opts);
            if (res.Status != PromptStatus.OK)
                throw new System.Exception("Selection canceled");

            return tr.GetObject(res.ObjectId, OpenMode.ForRead) as Entity;
        }

        private List<Entity> FindMatchingEntities(Transaction tr, Database db, Entity sample)
        {
            List<Entity> results = new List<Entity>();
            BlockTableRecord modelSpace = GetModelSpace(tr, db);

            foreach (ObjectId id in modelSpace)
            {
                Entity entity = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (entity != null)
                {
                    // Проверяем тип объекта и соответствие параметрам
                    if (entity is Circle circle && sample is Circle sampleCircle)
                    {
                        if (circle.Layer == sampleCircle.Layer &&
                            circle.Radius == sampleCircle.Radius &&
                            circle.Color.ColorValue.ToArgb() == sampleCircle.Color.ColorValue.ToArgb())
                        {
                            results.Add(circle);
                        }
                    }
                    else if (entity is Arc arc && sample is Arc sampleArc)
                    {
                        if (arc.Layer == sampleArc.Layer &&
                            arc.Radius == sampleArc.Radius &&
                            arc.Color.ColorValue.ToArgb() == sampleArc.Color.ColorValue.ToArgb())
                        {
                            results.Add(arc);
                        }
                    }
                    else if (entity is BlockReference block && sample is BlockReference sampleBlock)
                    {
                        if (block.Layer == sampleBlock.Layer &&
                            block.Name == sampleBlock.Name &&
                            block.Color.ColorValue.ToArgb() == sampleBlock.Color.ColorValue.ToArgb())
                        {
                            results.Add(block);
                        }
                    }
                }
            }
            return results;
        }

        private Point3d GetEntityCenter(Entity entity)
        {
            if (entity is Circle circle)
                return circle.Center;
            else if (entity is Arc arc)
                return arc.Center;
            else if (entity is BlockReference block)
                return block.Position; // Точка вставки блока
            else
                throw new ArgumentException("Entity must be Circle, Arc or BlockReference");
        }

        private double CalculateSearchRadius(Entity entity, double maxTextDimension)
        {
            if (entity is Circle circle)
            {
                return circle.Radius * 3 + maxTextDimension;
            }
            else if (entity is Arc arc)
            {
                return arc.Radius * 3 + maxTextDimension;
            }
            else if (entity is BlockReference block)
            {
                // Для блоков используем размер bounding box
                try
                {
                    Extents3d extents = block.GeometricExtents;
                    double width = extents.MaxPoint.X - extents.MinPoint.X;
                    double height = extents.MaxPoint.Y - extents.MinPoint.Y;
                    double diagonal = Math.Sqrt(width * width + height * height);
                    return diagonal * 1.5 + maxTextDimension;
                }
                catch
                {
                    // Если не удалось получить bounding box, используем фиксированный размер
                    return 100.0 + maxTextDimension;
                }
            }
            else
            {
                throw new ArgumentException("Unsupported entity type");
            }
        }

        private double CalculateMaxTextDimension(List<Entity> texts)
        {
            double maxDimension = 0;
            foreach (Entity text in texts)
            {
                try
                {
                    Extents3d bounds = GetTextBounds(text);
                    double width = bounds.MaxPoint.X - bounds.MinPoint.X;
                    double height = bounds.MaxPoint.Y - bounds.MinPoint.Y;
                    maxDimension = Math.Max(maxDimension, Math.Max(width, height));
                }
                catch
                {
                    // Игнорируем ошибки при вычислении размеров текста
                }
            }
            return maxDimension;
        }

        private double GetMinDistanceBetween(Entity text, Point3d center)
        {
            try
            {
                Extents3d bounds = GetTextBounds(text);

                // Вычисляем минимальное расстояние до ограничивающей рамки
                double dx = Math.Max(Math.Max(bounds.MinPoint.X - center.X, center.X - bounds.MaxPoint.X), 0);
                double dy = Math.Max(Math.Max(bounds.MinPoint.Y - center.Y, center.Y - bounds.MaxPoint.Y), 0);

                return Math.Sqrt(dx * dx + dy * dy);
            }
            catch
            {
                return GetTextPosition(text).DistanceTo(center);
            }
        }

        private Extents3d GetTextBounds(Entity text)
        {
            if (text is DBText dbText)
            {
                return dbText.GeometricExtents;
            }
            else if (text is MText mText)
            {
                return mText.GeometricExtents;
            }
            else
            {
                var pos = GetTextPosition(text);
                return new Extents3d(pos, pos);
            }
        }

        private Entity SelectText(Editor ed, Transaction tr)
        {
            var opts = new PromptEntityOptions("\nSelect sample text/MTEXT: ")
            {
                AllowNone = false,
                AllowObjectOnLockedLayer = true
            };
            opts.SetRejectMessage("Selected object is not text or MTEXT");
            opts.AddAllowedClass(typeof(DBText), true);
            opts.AddAllowedClass(typeof(MText), true);

            PromptEntityResult res = ed.GetEntity(opts);
            if (res.Status != PromptStatus.OK)
                throw new System.Exception("Selection canceled");

            return tr.GetObject(res.ObjectId, OpenMode.ForRead) as Entity;
        }

        private List<Entity> FindMatchingTexts(Transaction tr, Database db, Entity sample)
        {
            List<Entity> results = new List<Entity>();
            BlockTableRecord modelSpace = GetModelSpace(tr, db);

            // Получаем свойства образца текста
            string sampleLayer = sample.Layer;
            int sampleColor = sample.Color.ColorValue.ToArgb();
            string sampleStyle = GetTextStyleName(sample, tr);
            double sampleHeight = GetTextHeight(sample);

            foreach (ObjectId id in modelSpace)
            {
                if (!IsTextType(id)) continue;

                Entity text = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (text != null &&
                    text.Layer == sampleLayer &&
                    text.Color.ColorValue.ToArgb() == sampleColor)
                {
                    // Проверяем стиль и высоту текста
                    string textStyle = GetTextStyleName(text, tr);
                    double textHeight = GetTextHeight(text);

                    if (textStyle == sampleStyle && Math.Abs(textHeight - sampleHeight) < 0.001)
                    {
                        results.Add(text);
                    }
                }
            }
            return results;
        }

        private double GetTextHeight(Entity text)
        {
            if (text is DBText dbText)
                return dbText.Height;
            else if (text is MText mText)
                return mText.TextHeight;
            else
                return 0;
        }

        private void MoveTextToPoint(Entity text, Point3d newPos)
        {
            text.UpgradeOpen();
            try
            {
                if (text is DBText dbText)
                {
                    Vector3d offset = newPos - dbText.Position;
                    dbText.TransformBy(Matrix3d.Displacement(offset));
                }
                else if (text is MText mText)
                {
                    Vector3d offset = newPos - mText.Location;
                    mText.TransformBy(Matrix3d.Displacement(offset));
                }
            }
            finally
            {
                text.DowngradeOpen();
            }
        }

        #region Helper Methods
        private BlockTableRecord GetModelSpace(Transaction tr, Database db)
        {
            return tr.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db),
                OpenMode.ForRead
            ) as BlockTableRecord;
        }

        private bool IsTextType(ObjectId id)
        {
            return id.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(DBText))) ||
                   id.ObjectClass.IsDerivedFrom(RXClass.GetClass(typeof(MText)));
        }

        private string GetTextStyleName(Entity text, Transaction tr)
        {
            if (text is DBText dbText)
            {
                return ((TextStyleTableRecord)tr.GetObject(dbText.TextStyleId, OpenMode.ForRead)).Name;
            }
            if (text is MText mText)
            {
                return ((TextStyleTableRecord)tr.GetObject(mText.TextStyleId, OpenMode.ForRead)).Name;
            }
            throw new ArgumentException("Not a text object");
        }

        private Point3d GetTextPosition(Entity text)
        {
            if (text is DBText dbText)
            {
                return dbText.Position;
            }
            if (text is MText mText)
            {
                return mText.Location;
            }
            throw new ArgumentException("Invalid text type");
        }
        #endregion
    }
}