using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace TopoBuilder // Убедитесь, что это ваш правильный namespace
{
    public class TerrainGenerator
    {
        // Поля для хранения свойств образцового текста
        private string _targetLayer;
        private ObjectId _targetLayerId;
        private Color _targetColor;
        private bool _colorByLayer;

       
        /// Главная команда AutoCAD для создания 3D-модели рельефа и вспомогательной 3D-поверхности.
      
        [CommandMethod("TOPOMODEL")]
        public void CreateTerrainSolidFinalCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) { Application.ShowAlertDialog("No active document."); return; }
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ObjectId finalSolidId = ObjectId.Null;
            ObjectId surfaceMeshId = ObjectId.Null;

            try
            {
                // Шаг 1: Получение свойств (слой, цвет) с указанного пользователем образцового текста.
                if (!GetSampleTextProperties(ed))
                {
                    ed.WriteMessage("\nOperation cancelled by user or error selecting sample text properties.");
                    return;
                }

                // Шаг 2: Сбор 3D-точек. X,Y из положения текста, Z из содержимого текста.
                // Фильтрация по свойствам образца. Обработка дубликатов XY (выбирается макс. Z).
                List<Point3d> points3d = CollectValidTextPoints3dIterative(ed, db);
                ed.WriteMessage($"\nFound {points3d.Count} unique XY points with valid Z elevations.");
                if (points3d.Count < 3)
                {
                    ed.WriteMessage("\nLess than 3 unique points found. Cannot create a 3D solid or surface.");
                    return;
                }

                // Шаг 3: Построение триангуляции Делоне по 2D-проекциям точек.
                List<int> triangleIndices = DelaunayTriangulation(points3d, ed);
                if (triangleIndices == null || triangleIndices.Count == 0)
                {
                    ed.WriteMessage("\nDelaunay triangulation failed or yielded no triangles.");
                    return;
                }
                ed.WriteMessage($"\nDelaunay triangulation resulted in {triangleIndices.Count / 3} triangles.");

                // Начало основной транзакции для создания геометрии
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Шаг 4 (Опционально, для визуализации): Создание PolyFaceMesh, представляющей 3D-топоповерхность.
                        ed.WriteMessage("\nCreating 3D PolyFaceMesh surface for visualization...");
                        surfaceMeshId = CreatePolyFaceMeshSurface(points3d, triangleIndices, db, tr, ed, "Topo_Surface_Visualization");
                        if (surfaceMeshId.IsNull || surfaceMeshId.IsErased)
                        {
                            ed.WriteMessage("\nWarning: Failed to create the 3D PolyFaceMesh surface for visualization.");
                        }
                        else
                        {
                            ed.WriteMessage($"\n3D PolyFaceMesh surface created (ObjectId: {surfaceMeshId}).");
                        }

                        // Шаг 5: Создание итогового 3D-солида на основе триангуляции.
                        // Каждому треугольнику соответствует призма с основанием на Z=0 и наклонным верхом.
                        ed.WriteMessage("\nCreating 3D Solid from triangulated prisms...");
                        finalSolidId = CreateSolidFromTriangles(points3d, triangleIndices, db, tr, ed);

                        // Завершение транзакции
                        if (finalSolidId == ObjectId.Null || finalSolidId.IsErased)
                        {
                            ed.WriteMessage("\nFailed to create the final 3D Solid object.");
                            tr.Abort(); // Откатываем все изменения, если солид не создан
                            ed.WriteMessage("\nTransaction aborted due to solid creation failure.");
                        }
                        else
                        {
                            ed.WriteMessage($"\nFinal 3D Solid successfully created (ObjectId: {finalSolidId}).");
                            tr.Commit();
                            ed.WriteMessage("\nTransaction committed. Solid and visualization surface (if created) are saved.");
                        }
                    }
                    catch (System.Exception exBuild)
                    {
                        tr.Abort(); // Откат транзакции при любой ошибке на этапе построения
                        ed.WriteMessage($"\nERROR during 3D geometry creation: {exBuild.GetType().Name} - {exBuild.Message}\nStackTrace: {exBuild.StackTrace}");
                    }
                }
            }
            catch (System.Exception exMain)
            {
                ed.WriteMessage($"\nUNEXPECTED ERROR in TOPOMODEL command: {exMain.GetType().Name} - {exMain.Message}\nStackTrace: {exMain.StackTrace}");
            }
        }

        /// Запрашивает у пользователя выбор образцового текстового объекта (DBText или MText)
        /// и считывает его свойства (слой и цвет) для дальнейшей фильтрации.
        private bool GetSampleTextProperties(Editor ed)
        {
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect sample text object (TEXT or MTEXT) to define properties for point selection: ");
            peo.SetRejectMessage("\nInvalid selection: Selected object is not a TEXT or MTEXT entity.");
            peo.AddAllowedClass(typeof(DBText), true);
            peo.AddAllowedClass(typeof(MText), true);
            PromptEntityResult result = ed.GetEntity(peo);

            if (result.Status != PromptStatus.OK) return false;

            using (Transaction tr = ed.Document.TransactionManager.StartTransaction())
            {
                Entity ent = tr.GetObject(result.ObjectId, OpenMode.ForRead) as Entity;
                if (ent == null) { tr.Abort(); return false; } // Не должно случиться, если GetEntity ОК

                _targetLayer = ent.Layer;
                _targetLayerId = ent.LayerId;
                _colorByLayer = ent.Color.ColorMethod == ColorMethod.ByLayer;

                if (_colorByLayer)
                {
                    LayerTableRecord ltr = tr.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                    _targetColor = ltr?.Color ?? Color.FromColorIndex(ColorMethod.ByAci, 7); // Белый по умолчанию, если цвет слоя не определен
                }
                else
                {
                    _targetColor = ent.Color;
                }
                ed.WriteMessage($"\nProperties for point selection set from sample:\n- Layer: '{_targetLayer}'\n- Color: {_targetColor} (Method: {(_colorByLayer ? "ByLayer" : "Explicit")})");
                tr.Commit(); // Транзакция только для чтения свойств
            }
            return true;
        }
              
        /// Сканирует пространство модели на наличие текстовых объектов (DBText, MText),
        /// соответствующих заданным свойствам слоя и цвета. Извлекает 3D-точки,
        /// где XY - из положения текста, а Z - из его числового содержимого.
        /// При наличии текстов с одинаковыми XY-координатами выбирается тот, у которого Z-значение максимально.

        private List<Point3d> CollectValidTextPoints3dIterative(Editor ed, Database db)
        {
            var rawPoints = new List<Point3d>();
            int totalTextScanned = 0, propertyMatchCount = 0, parseErrorCount = 0;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord ms = tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                if (ms == null) { ed.WriteMessage("\nCritical Error: Model space not found."); return rawPoints; }

                ed.WriteMessage($"\nScanning ModelSpace for TEXT/MTEXT objects on layer '{_targetLayer}' with specified color...");
                foreach (ObjectId objId in ms)
                {
                    // Быстрая проверка типа объекта
                    if (!objId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(DBText))) &&
                        !objId.ObjectClass.IsDerivedFrom(RXObject.GetClass(typeof(MText)))) continue;

                    Entity ent = tr.GetObject(objId, OpenMode.ForRead, false, true) as Entity; // openErased = false, forceOpen = true
                    if (ent == null) continue;

                    totalTextScanned++;

                    if (ent.LayerId != _targetLayerId) continue;

                    Color entityColor = ent.Color;
                    if (ent.Color.ColorMethod == ColorMethod.ByLayer)
                    {
                        LayerTableRecord ltr = tr.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                        entityColor = ltr?.Color ?? Color.FromColorIndex(ColorMethod.ByAci, 7); // Белый, если цвет слоя не задан
                    }

                    if (!entityColor.ColorValue.Equals(_targetColor.ColorValue)) continue;

                    propertyMatchCount++;
                    Point3d textPosition = Point3d.Origin;
                    string textContentValue = "";

                    if (ent is DBText dbText) { textPosition = dbText.Position; textContentValue = dbText.TextString; }
                    else if (ent is MText mText) { textPosition = mText.Location; textContentValue = mText.Contents; }

                    if (double.TryParse(textContentValue.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double zValue))
                    {
                        rawPoints.Add(new Point3d(textPosition.X, textPosition.Y, zValue));
                    }
                    else { parseErrorCount++; }
                }
                ed.WriteMessage($"\nText scan results:\n- Total TEXT/MTEXT entities found: {totalTextScanned}\n- Matched layer and color: {propertyMatchCount}\n- Errors parsing Z-value from text: {parseErrorCount}");
                tr.Commit(); // Транзакция только для чтения
            }

            if (rawPoints.Count == 0) return rawPoints;

            ed.WriteMessage($"\nFiltering duplicate XY coordinates (keeping max Z for each XY)...");
            var uniquePoints = rawPoints.GroupBy(p => new Point2d(p.X, p.Y), new Point2dComparer(0.001)) // Толерантность 1мм для группировки XY
                                      .Select(g => g.OrderByDescending(p => p.Z).First()) // Выбираем точку с макс. Z
                                      .ToList();
            if (rawPoints.Count != uniquePoints.Count)
                ed.WriteMessage($"  {rawPoints.Count - uniquePoints.Count} duplicate XY entries consolidated.");

            return uniquePoints;
        }


        /// Создает и добавляет в чертеж объект PolyFaceMesh, представляющий 3D-поверхность,
        /// на основе списка 3D-вершин и списка индексов треугольников.

        private ObjectId CreatePolyFaceMeshSurface(List<Point3d> points3d, List<int> triangleIndices, Database db, Transaction tr, Editor ed, string layerName)
        {
            ObjectId meshId = ObjectId.Null;
            PolyFaceMesh pfaceMeshForWrite = null;

            try
            {
                using (PolyFaceMesh pfaceMeshInMemory = new PolyFaceMesh())
                {
                    pfaceMeshInMemory.SetDatabaseDefaults(db);
                    LayerTable lt = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    if (lt == null) { ed.WriteMessage("\nError: Cannot access LayerTable for PolyFaceMesh."); return ObjectId.Null; }

                    if (!lt.Has(layerName))
                    {
                        try
                        {
                            lt.UpgradeOpen();
                            LayerTableRecord ltrNew = new LayerTableRecord { Name = layerName };
                            lt.Add(ltrNew); tr.AddNewlyCreatedDBObject(ltrNew, true);
                            pfaceMeshInMemory.Layer = layerName;
                            lt.DowngradeOpen();
                        }
                        catch (System.Exception exLayer) { ed.WriteMessage($"\nWarning: Could not create layer '{layerName}' for PolyFaceMesh: {exLayer.Message}. Using current layer."); }
                    }
                    else { pfaceMeshInMemory.Layer = layerName; }
                    pfaceMeshInMemory.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256); // Цвет по слою

                    BlockTableRecord ms = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite) as BlockTableRecord;
                    if (ms == null) { ed.WriteMessage("\nError: Cannot access ModelSpace for PolyFaceMesh."); return ObjectId.Null; }

                    meshId = ms.AppendEntity(pfaceMeshInMemory);
                    tr.AddNewlyCreatedDBObject(pfaceMeshInMemory, true);

                    pfaceMeshForWrite = tr.GetObject(meshId, OpenMode.ForWrite) as PolyFaceMesh;
                    if (pfaceMeshForWrite == null) { ed.WriteMessage("\nError: Could not open created PolyFaceMesh for writing."); return ObjectId.Null; }
                }

                foreach (Point3d pt in points3d)
                { pfaceMeshForWrite.AppendVertex(new PolyFaceMeshVertex(pt)); }

                int faceCount = 0;
                for (int i = 0; i < triangleIndices.Count; i += 3)
                {
                    pfaceMeshForWrite.AppendFaceRecord(new FaceRecord((short)(triangleIndices[i] + 1),
                                                                    (short)(triangleIndices[i + 1] + 1),
                                                                    (short)(triangleIndices[i + 2] + 1), 0));
                    faceCount++;
                }

                if (faceCount > 0)
                {
                    // ed.WriteMessage($"\nPolyFaceMesh surface populated with {points3d.Count} vertices and {faceCount} faces (ID: {meshId}).");
                    return meshId;
                }
                else
                {
                    ed.WriteMessage("\nNo faces were generated for PolyFaceMesh. Erasing empty mesh entity.");
                    if (pfaceMeshForWrite != null && !pfaceMeshForWrite.IsDisposed) { pfaceMeshForWrite.Erase(); }
                    return ObjectId.Null;
                }
            }
            catch (System.Exception ex) { ed.WriteMessage($"\nEXCEPTION in CreatePolyFaceMeshSurface: {ex.GetType().Name} - {ex.Message}\n{ex.StackTrace}"); return ObjectId.Null; }
        }


        /// Создает 3D-солид. Для каждого треугольника из триангуляции создается призматическое тело:
        /// основание на Z=0 (проекция треугольника), которое выдавливается вверх,
        /// а затем отсекается плоскостью, проходящей через исходные 3D-вершины треугольника.
        /// Все созданные призмы затем объединяются в одно тело.

        private ObjectId CreateSolidFromTriangles(List<Point3d> vertices, List<int> triangleIndices, Database db, Transaction tr, Editor ed)
        {
            ed.WriteMessage("\n--- Starting creation of 3D solid from individual prisms (Extrude+Slice method)...");
            Solid3d finalSolid = null;
            List<Solid3d> solidParts = new List<Solid3d>();

            // Смещение основания. 0.0 для идеально ровных призм (могут быть проблемы с объединением).
            // Малое значение (напр. 0.0001) может улучшить объединение без заметных искажений.
            double baseOffsetAmount = 0.0; // Установлено в 0.0, так как это дало лучшие отдельные призмы

            try
            {
                BlockTableRecord ms = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                if (ms == null) throw new InvalidOperationException("Cannot get ModelSpace for writing solid parts.");

                int successfullyCreatedParts = 0;
                for (int i = 0; i < triangleIndices.Count; i += 3)
                {
                    int idx1 = triangleIndices[i], idx2 = triangleIndices[i + 1], idx3 = triangleIndices[i + 2];
                    if (idx1 < 0 || idx1 >= vertices.Count || idx2 < 0 || idx2 >= vertices.Count || idx3 < 0 || idx3 >= vertices.Count)
                    { ed.WriteMessage($"\n   - SKIPPING (Triangle {i / 3}): Invalid vertex index from triangulation."); continue; }

                    Point3d p1_top = vertices[idx1]; Point3d p2_top = vertices[idx2]; Point3d p3_top = vertices[idx3];

                    Vector3d vec_a = p2_top - p1_top; Vector3d vec_b = p3_top - p1_top;
                    Vector3d trianglePlaneNormal = vec_a.CrossProduct(vec_b);
                    if (trianglePlaneNormal.Length < 1e-9) // Проверка на вырожденный (коллинеарный) треугольник
                    { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Degenerate top triangle (vertices are collinear)."); continue; }

                    Region bottomRegion = null;
                    Solid3d extrudedPrism = null;
                    List<DBObject> tempGeometryForDisposalInLoop = new List<DBObject>();

                    try
                    {
                        // Шаг 1: Создание 2D-региона основания на Z=0
                        Point2d pt1_base_2d = new Point2d(p1_top.X, p1_top.Y);
                        Point2d pt2_base_2d = new Point2d(p2_top.X, p2_top.Y);
                        Point2d pt3_base_2d = new Point2d(p3_top.X, p3_top.Y);

                        Polyline basePolyline2d = new Polyline();
                        basePolyline2d.AddVertexAt(0, pt1_base_2d, 0, 0, 0);
                        basePolyline2d.AddVertexAt(1, pt2_base_2d, 0, 0, 0);
                        basePolyline2d.AddVertexAt(2, pt3_base_2d, 0, 0, 0);
                        basePolyline2d.Closed = true;
                        tempGeometryForDisposalInLoop.Add(basePolyline2d);

                        DBObject curveForBottomRegion = basePolyline2d;
                        // Логика смещения (baseOffsetAmount) намеренно удалена для "ровных" призм,
                        // так как это дало наилучший результат для отдельных частей.
                        // Если объединение все еще проблемное, можно вернуть сюда код для baseOffsetAmount > 0.

                        using (DBObjectCollection regionSourceBottom = new DBObjectCollection { curveForBottomRegion })
                        {
                            DBObjectCollection regionsBottomCol = null;
                            try { regionsBottomCol = Region.CreateFromCurves(regionSourceBottom); }
                            catch (System.Exception exReg) { ed.WriteMessage($"     - Exception creating bottom region for T{i / 3}: {exReg.Message}"); }

                            if (regionsBottomCol != null && regionsBottomCol.Count == 1 && regionsBottomCol[0] is Region)
                            { bottomRegion = (regionsBottomCol[0] as Region).Clone() as Region; }

                            if (regionsBottomCol != null) { foreach (DBObject r_obj in regionsBottomCol) { if (r_obj != bottomRegion && r_obj != null && !r_obj.IsDisposed) r_obj.Dispose(); } }
                        }

                        if (bottomRegion == null || bottomRegion.IsDisposed) { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Bottom Region is null or disposed after creation attempt."); continue; }
                        if (bottomRegion.Area < 1e-7) { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Bottom Region area is too small ({bottomRegion.Area:E1})."); bottomRegion.Dispose(); continue; }
                        tempGeometryForDisposalInLoop.Add(bottomRegion);

                        // Шаг 2: Экструзия основания вверх
                        extrudedPrism = new Solid3d();
                        extrudedPrism.SetDatabaseDefaults();
                        double maxZ_top = Math.Max(p1_top.Z, Math.Max(p2_top.Z, p3_top.Z));
                        double extrusionHeight = (maxZ_top >= 0 ? maxZ_top : 0) + 0.5; // Буфер 0.5 единиц над макс. точкой
                        if (extrusionHeight < 0.1) extrusionHeight = 0.1; // Минимальная высота

                        extrudedPrism.Extrude(bottomRegion, extrusionHeight, 0.0);
                        ms.AppendEntity(extrudedPrism);
                        tr.AddNewlyCreatedDBObject(extrudedPrism, true);

                        if (extrudedPrism.IsNull) { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Extruded prism is IsNull."); extrudedPrism.Dispose(); continue; } // Dispose C# wrapper
                        if (!ValidateSolid(extrudedPrism, ed, $"Extruded Prism (T{i / 3})")) { continue; }

                        Extents3d prismExtents = extrudedPrism.GeometricExtents;
                        if (Math.Abs(prismExtents.MinPoint.Z) > 0.001)
                        {
                            extrudedPrism.TransformBy(Matrix3d.Displacement(Vector3d.ZAxis * -prismExtents.MinPoint.Z));
                            if (Math.Abs(extrudedPrism.GeometricExtents.MinPoint.Z) > 0.001) { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Base Z correction failed for extruded prism."); continue; }
                        }

                        // Шаг 3: Отсечение (Slice) для формирования наклонной верхней грани
                        Point3d planeOrigin = p1_top;
                        Vector3d sliceNormalForCuttingTop;
                        // Нормаль плоскости сечения должна быть направлена ВНИЗ, чтобы с флагом keepSide=true осталась НИЖНЯЯ часть.
                        if (trianglePlaneNormal.Z >= 0) { sliceNormalForCuttingTop = trianglePlaneNormal.Negate(); } // Инвертируем, если смотрела вверх
                        else { sliceNormalForCuttingTop = trianglePlaneNormal; } // Уже смотрит вниз

                        if (sliceNormalForCuttingTop.IsZeroLength()) { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Slice normal is zero. Original normal Z was {trianglePlaneNormal.Z:F2}."); continue; }

                        Plane slicePlane = new Plane(planeOrigin, sliceNormalForCuttingTop);
                        bool keepSide = true;

                        extrudedPrism.Slice(slicePlane, keepSide);

                        if (extrudedPrism.IsNull) { ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Sliced prism became IsNull."); extrudedPrism.Dispose(); continue; } // Dispose C# wrapper
                        if (!ValidateSolid(extrudedPrism, ed, $"Sliced Solid (T{i / 3})")) { continue; }

                        Extents3d slicedExtents = extrudedPrism.GeometricExtents;
                        // Проверка Z-координат после отсечения (с допусками)
                        bool zCheckOk = true;
                        if (Math.Abs(slicedExtents.MinPoint.Z) > 0.01) { zCheckOk = false; }
                        if (zCheckOk && (slicedExtents.MaxPoint.Z > maxZ_top + 0.01 || slicedExtents.MaxPoint.Z < Math.Min(p1_top.Z, Math.Min(p2_top.Z, p3_top.Z)) - 0.01)) { zCheckOk = false; }

                        if (!zCheckOk)
                        {
                            ed.WriteMessage($"   - SKIPPING (Triangle {i / 3}): Failed Z extents check after slice. MinZ:{slicedExtents.MinPoint.Z:F3}, MaxZ:{slicedExtents.MaxPoint.Z:F3}, Expected Top ~{maxZ_top:F3}. Solid ID: {extrudedPrism.ObjectId}");
                            // Оставляем некорректный солид в чертеже для отладки (он будет удален при откате транзакции, если финальный солид не создан)
                            continue;
                        }

                        solidParts.Add(extrudedPrism);
                        successfullyCreatedParts++;
                    }
                    catch (System.Exception exLoopCaught)
                    { // Изменено имя переменной
                        ed.WriteMessage($"\n   - ERROR (Triangle {i / 3}) processing: {exLoopCaught.GetType().Name} - {exLoopCaught.Message}.");
                        if (extrudedPrism != null && !extrudedPrism.IsDisposed && !solidParts.Contains(extrudedPrism))
                        {
                            if (!extrudedPrism.ObjectId.IsNull && !extrudedPrism.ObjectId.IsErased) try { extrudedPrism.Erase(); } catch { }
                            extrudedPrism.Dispose();
                        }
                    }
                    finally
                    {
                        foreach (DBObject tempObj in tempGeometryForDisposalInLoop)
                        { if (tempObj != null && !tempObj.IsDisposed) tempObj.Dispose(); }
                    }
                }

                ed.WriteMessage($"\n--- Individual prism creation attempts finished ---");
                ed.WriteMessage($"Total prisms attempted: {triangleIndices.Count / 3}");
                ed.WriteMessage($"Successfully created individual prisms: {successfullyCreatedParts}");

                if (solidParts.Count == 0) { ed.WriteMessage("\nNo valid prisms were created. Final solid cannot be formed."); return ObjectId.Null; }

                // Шаг 4: Объединение всех успешно созданных призматических частей
                ed.WriteMessage($"\nUnioning {solidParts.Count} prism(s)...");

                finalSolid = solidParts[0]; // Первая часть уже в базе данных

                for (int j = 1; j < solidParts.Count; j++)
                {
                    Solid3d nextPart = solidParts[j]; // Эта часть также уже в базе
                    if (finalSolid.IsDisposed || nextPart.IsDisposed)
                    {
                        ed.WriteMessage($"\n   - SKIPPING UNION: A solid part is disposed (finalSolid or part index {j}).");
                        continue;
                    }
                    try
                    {
                        finalSolid.BooleanOperation(BooleanOperationType.BoolUnite, nextPart);
                        Entity dbEntityNextPart = tr.GetObject(nextPart.ObjectId, OpenMode.ForWrite, false, true) as Entity;
                        dbEntityNextPart?.Erase(); // Удаляем представление объединенной части из базы
                    }
                    catch (System.Exception exBoolUnion)
                    { // Изменено имя переменной
                        ed.WriteMessage($"\n   - ERROR unioning with part (DB ID: {nextPart.ObjectId}): {exBoolUnion.Message}. This part might remain separate.");
                        // Оставляем nextPart в чертеже, если объединение не удалось, для анализа.
                    }
                }
                ed.WriteMessage($"\nUnion operation finished.");

                // Освобождаем C# обертки для всех частей, которые были использованы/объединены, кроме финальной.
                foreach (var part in solidParts)
                {
                    if (part != finalSolid && part != null && !part.IsDisposed) { part.Dispose(); }
                }

                if (finalSolid == null || finalSolid.IsDisposed || finalSolid.ObjectId.IsNull || finalSolid.ObjectId.IsErased)
                {
                    ed.WriteMessage("\nError: Final solid is invalid or was not properly created/unioned after union process.");
                    if (finalSolid != null && !finalSolid.IsDisposed && finalSolid.ObjectId != ObjectId.Null && !finalSolid.ObjectId.IsErased)
                    {
                        try { (tr.GetObject(finalSolid.ObjectId, OpenMode.ForWrite) as Entity)?.Erase(); } catch { }
                    }
                    finalSolid?.Dispose(); // Освобождаем C# обертку, если она еще существует
                    return ObjectId.Null;
                }

                if (!ValidateSolid(finalSolid, ed, "Final Unioned Solid"))
                {
                    ed.WriteMessage("\nWarning: Final unioned solid may be invalid or have zero volume.");
                }
                return finalSolid.ObjectId;
            }
            catch (System.Exception exMainCaught)
            { // Изменено имя переменной
                ed.WriteMessage($"\n--- FATAL ERROR in CreateSolidFromTriangles: {exMainCaught.GetType().Name} - {exMainCaught.Message}\n{exMainCaught.StackTrace}");
                foreach (var part in solidParts) { if (part != null && !part.IsDisposed) part.Dispose(); }
                if (finalSolid != null && !finalSolid.IsDisposed) finalSolid.Dispose();
                throw; // Перебрасываем, чтобы внешняя транзакция (в CreateTerrainSolidFinalCommand) откатилась
            }
        }

        // --- Вспомогательные методы для триангуляции и валидации ---
        // (Эти методы остаются такими же, как были в вашем рабочем коде)


        /// Проверяет валидность 3D-тела (не null, не удалено, не "пустое" и имеет минимальный объем).

        private bool ValidateSolid(Solid3d solid, Editor ed, string context, double minVolume = 1e-7)
        {
            if (solid == null || solid.IsDisposed || solid.IsNull)
            { ed.WriteMessage($"\n   - VALIDATION FAILED ({context}): Solid is null, disposed, or IsNull state."); return false; }
            try
            {
                double volume;
                // Для получения MassProperties солид должен быть в базе или должна быть активна транзакция.
                if (solid.ObjectId == ObjectId.Null || solid.Database == null)
                {
                    if (solid.GeometricExtents == null) { ed.WriteMessage($"\n   - VALIDATION FAILED ({context}): Solid (not in DB) has NULL GeometricExtents."); return false; }
                    // Для солидов не в базе, если они не IsNull и имеют экстенты, считаем условно валидными.
                    // Объемную проверку для них здесь пропускаем.
                    return true;
                }
                else
                {
                    volume = solid.MassProperties.Volume;
                }
                if (volume < minVolume) { ed.WriteMessage($"\n   - VALIDATION FAILED ({context}): Solid volume ({volume:E3}) is less than minVolume ({minVolume:E3})."); return false; }
                return true;
            }
            catch (Autodesk.AutoCAD.Runtime.Exception acadEx)
            {
                ed.WriteMessage($"\n   - VALIDATION AutoCAD ERROR ({context}) for MassProperties: {acadEx.Message} (Status: {acadEx.ErrorStatus})");
                return false;
            }
            catch (System.Exception ex) { ed.WriteMessage($"\n   - VALIDATION SYSTEM ERROR ({context}) for MassProperties: {ex.Message}"); return false; }
        }


        /// Выполняет триангуляцию Делоне для набора 2D-проекций 3D-точек.
        /// Возвращает список индексов вершин, образующих треугольники.

        private List<int> DelaunayTriangulation(List<Point3d> points3d, Editor ed)
        {
            var points2d = points3d.Select(p => new Point2d(p.X, p.Y)).ToList();
            int originalPointCount = points2d.Count;
            if (originalPointCount < 3) { ed.WriteMessage("\nNot enough unique points for Delaunay triangulation (minimum 3 required)."); return new List<int>(); }

            var superTriangleVertices = CreateSuperTriangle(points2d);
            points2d.AddRange(superTriangleVertices);

            int stv1Idx = originalPointCount; int stv2Idx = originalPointCount + 1; int stv3Idx = originalPointCount + 2;
            var triangles = new List<int> { stv1Idx, stv2Idx, stv3Idx }; // Начинаем с супер-треугольника

            for (int i = 0; i < originalPointCount; i++)
            { // Итерация по каждой исходной точке
                Point2d currentPoint = points2d[i];
                var badTrianglesIndicesInList = new List<int>();
                var polygonEdges = new List<Edge>();

                // Находим все треугольники, чья описанная окружность содержит текущую точку
                for (int j = 0; j < triangles.Count; j += 3)
                {
                    int idx1 = triangles[j], idx2 = triangles[j + 1], idx3 = triangles[j + 2];
                    Point2d p1_2d = points2d[idx1], p2_2d = points2d[idx2], p3_2d = points2d[idx3];
                    if (IsPointInCircumcircle(currentPoint, p1_2d, p2_2d, p3_2d))
                    {
                        badTrianglesIndicesInList.Add(j);
                        polygonEdges.Add(new Edge(idx1, idx2)); polygonEdges.Add(new Edge(idx2, idx3)); polygonEdges.Add(new Edge(idx3, idx1));
                    }
                }
                // Удаляем "плохие" треугольники
                for (int k = badTrianglesIndicesInList.Count - 1; k >= 0; k--) { triangles.RemoveRange(badTrianglesIndicesInList[k], 3); }

                // Находим уникальные ребра, образующие "дыру"
                var boundaryEdges = polygonEdges.GroupBy(e => e).Where(g => g.Count() == 1).Select(g => g.Key).ToList();

                // Создаем новые треугольники, соединяя ребра "дыры" с текущей точкой
                foreach (var edge in boundaryEdges) { triangles.Add(edge.A); triangles.Add(edge.B); triangles.Add(i); }
            }

            // Формируем финальный список треугольников, удаляя те, что включают вершины супер-треугольника
            var finalTriangles = new List<int>();
            for (int i = 0; i < triangles.Count; i += 3)
            {
                int v1 = triangles[i], v2 = triangles[i + 1], v3 = triangles[i + 2];
                if (v1 < originalPointCount && v2 < originalPointCount && v3 < originalPointCount)
                { finalTriangles.Add(v1); finalTriangles.Add(v2); finalTriangles.Add(v3); }
            }
            return finalTriangles;
        }


        /// Создает "супер-треугольник", который гарантированно содержит все исходные точки.
        /// Используется как начальная структура в алгоритме триангуляции Делоне.
        private List<Point2d> CreateSuperTriangle(List<Point2d> points)
        {
            if (points == null || points.Count == 0) return new List<Point2d>(); // Защита от пустого списка
            double minX = points[0].X, maxX = points[0].X, minY = points[0].Y, maxY = points[0].Y;
            foreach (var p in points.Skip(1))
            {
                if (p.X < minX) minX = p.X; if (p.X > maxX) maxX = p.X;
                if (p.Y < minY) minY = p.Y; if (p.Y > maxY) maxY = p.Y;
            }
            double dx = maxX - minX, dy = maxY - minY;
            double deltaMax = Math.Max(dx, dy);
            if (deltaMax < 1e-6) deltaMax = Math.Max(Math.Abs(minX), Math.Max(Math.Abs(maxX), Math.Max(Math.Abs(minY), Math.Abs(maxY)))) + 1.0; // Если все точки совпадают
            if (deltaMax < 1e-6) deltaMax = 100.0; // Абсолютный минимум, если все точки (0,0)

            double midX = (minX + maxX) / 2.0; double midY = (minY + maxY) / 2.0;

            // Вершины супер-треугольника должны быть достаточно далеко
            return new List<Point2d> {
                new Point2d(midX - 20 * deltaMax, midY - 10 * deltaMax),
                new Point2d(midX + 20 * deltaMax, midY - 10 * deltaMax),
                new Point2d(midX                 , midY + 20 * deltaMax)
            };
        }

        /// Проверяет, находится ли точка p внутри описанной окружности треугольника abc.
        private bool IsPointInCircumcircle(Point2d p, Point2d a, Point2d b, Point2d c)
        {
            // Используем тест на основе определителя.
            // Переносим начало координат в точку p для улучшения точности вычислений.
            double ax = a.X - p.X, ay = a.Y - p.Y;
            double bx = b.X - p.X, by = b.Y - p.Y;
            double cx = c.X - p.X, cy = c.Y - p.Y;

            double aSq = ax * ax + ay * ay;
            double bSq = bx * bx + by * by;
            double cSq = cx * cx + cy * cy;

            double det = ax * (by * cSq - cy * bSq) -
                         ay * (bx * cSq - cx * bSq) +
                         aSq * (bx * cy - cx * by);

            // Знак определителя зависит от ориентации вершин a, b, c.
            // Определяем ориентацию (порядок обхода) вершин a, b, c.
            double orientation = (b.X - a.X) * (c.Y - a.Y) - (b.Y - a.Y) * (c.X - a.X);
            double epsilon = 1e-9; // Малая величина для сравнения с нулем

            if (Math.Abs(orientation) < epsilon) return false; // Точки a, b, c коллинеарны, окружность не определена.

            // Если ориентация (a,b,c) против часовой стрелки, точка p внутри, если det > 0.
            // Если ориентация (a,b,c) по часовой стрелке, точка p внутри, если det < 0.
            if (orientation > 0) return det > epsilon;
            else return det < -epsilon;
        }

        /// Структура для представления ребра между двумя вершинами (по их индексам).
        /// Индексы хранятся в каноническом порядке (меньший, затем больший) для упрощения сравнения.
        private struct Edge : IEquatable<Edge>
        {
            public int A { get; }
            public int B { get; }
            public Edge(int a, int b) { A = Math.Min(a, b); B = Math.Max(a, b); }
            public bool Equals(Edge other) => A == other.A && B == other.B;
            public override bool Equals(object obj) => obj is Edge edge && Equals(edge);
            public override int GetHashCode() => A.GetHashCode() ^ (B.GetHashCode() * 397); // Комбинация хеш-кодов
        }

        /// Класс для сравнения объектов Point2d с заданной точностью (толерантностью).
        /// Используется для группировки близких точек при удалении дубликатов.
        private class Point2dComparer : IEqualityComparer<Point2d>
        {
            private readonly double _tolerance;
            private readonly double _invTolerance; // Для GetHashCode
            public Point2dComparer(double tolerance)
            { _tolerance = Math.Max(tolerance, 1e-9); _invTolerance = 1.0 / _tolerance; }

            public bool Equals(Point2d a, Point2d b) =>
                Math.Abs(a.X - b.X) < _tolerance && Math.Abs(a.Y - b.Y) < _tolerance;

            public int GetHashCode(Point2d obj)
            {
                // Квантуем координаты для группировки близких точек
                int hashX = Math.Round(obj.X * _invTolerance).GetHashCode();
                int hashY = Math.Round(obj.Y * _invTolerance).GetHashCode();
                return hashX ^ (hashY * 397); // Простое XOR-сочетание хеш-кодов
            }
        }

    } // Конец класса TerrainGenerator
} // Конец namespace TopoBuilder