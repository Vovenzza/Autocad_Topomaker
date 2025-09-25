using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace TopoBuilder
{
    public class TopoCommands
    {
        public static List<Point3d> GeneratedTerrainPoints { get; } = new List<Point3d>();

        [CommandMethod("TOPOMODEL")]
        public void CreateTopographyPoints()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            GeneratedTerrainPoints.Clear();
            ed.WriteMessage("\nInitializing topography point generation...");

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    Entity sample = GetSampleEntity(ed, tr);
                    if (sample == null) return;

                    short targetColorIndex = GetEffectiveColorIndex(sample, tr);

                    TypedValue[] filterValues = {
                        new TypedValue((int)DxfCode.Start, "TEXT,MTEXT"),
                        new TypedValue((int)DxfCode.LayerName, sample.Layer)
                    };

                    PromptSelectionResult selection = ed.SelectAll(new SelectionFilter(filterValues));
                    if (!ValidateSelection(ed, selection)) return;

                    ProcessEntities(tr, ed, selection.Value, db, targetColorIndex);

                    tr.Commit();
                    ed.Regen();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nERROR: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        private Entity GetSampleEntity(Editor ed, Transaction tr)
        {
            var options = new PromptEntityOptions("\nSelect sample height mark:")
            {
                AllowNone = false,
                AllowObjectOnLockedLayer = true
            };
            options.SetRejectMessage("\nInvalid selection. Select Text/MText");
            options.AddAllowedClass(typeof(DBText), true);
            options.AddAllowedClass(typeof(MText), true);

            PromptEntityResult result = ed.GetEntity(options);
            return result.Status == PromptStatus.OK
                ? tr.GetObject(result.ObjectId, OpenMode.ForRead) as Entity
                : null;
        }

        private bool ValidateSelection(Editor ed, PromptSelectionResult result)
        {
            if (result.Status != PromptStatus.OK || result.Value.Count == 0)
            {
                ed.WriteMessage("\nNo TEXT or MTEXT found on selected layer");
                return false;
            }
            return true;
        }

        private short GetEffectiveColorIndex(Entity entity, Transaction tr)
        {
            if (entity.Color.ColorMethod == ColorMethod.ByLayer)
            {
                LayerTableRecord ltr = tr.GetObject(entity.LayerId, OpenMode.ForRead) as LayerTableRecord;
                return (short)ltr.Color.ColorIndex;
            }
            return (short)entity.Color.ColorIndex;
        }

        private void ProcessEntities(
            Transaction tr,
            Editor ed,
            SelectionSet selection,
            Database db,
            short targetColorIndex)
        {
            BlockTableRecord ms = tr.GetObject(
                SymbolUtilityServices.GetBlockModelSpaceId(db),
                OpenMode.ForWrite) as BlockTableRecord;

            double tolerance = db.Insunits == UnitsValue.Millimeters ? 0.1 : 0.001;
            var pointMap = new Dictionary<Point3d, double>(new Point2dComparer(tolerance));

            int processed = 0, errors = 0, colorMismatch = 0;

            foreach (SelectedObject obj in selection)
            {
                try
                {
                    Entity ent = tr.GetObject(obj.ObjectId, OpenMode.ForRead) as Entity;
                    short entColor = GetEffectiveColorIndex(ent, tr);

                    if (entColor != targetColorIndex)
                    {
                        colorMismatch++;
                        continue;
                    }

                    if (ProcessEntity(ent, pointMap, tolerance))
                        processed++;
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError processing entity: {ex.Message}");
                    errors++;
                }
            }

            // Add filtered points to drawing
            foreach (var entry in pointMap)
            {
                using (DBPoint dbPoint = new DBPoint(new Point3d(entry.Key.X, entry.Key.Y, entry.Value)))
                {
                    ms.AppendEntity(dbPoint);
                    tr.AddNewlyCreatedDBObject(dbPoint, true);
                    GeneratedTerrainPoints.Add(dbPoint.Position);
                }
            }

            ed.WriteMessage(
                $"\nResults: {processed} points processed | " +
                $"{pointMap.Count} unique points added | " +
                $"{colorMismatch} color mismatches | " +
                $"{errors} errors"
            );
        }

        private bool ProcessEntity(
            Entity ent,
            Dictionary<Point3d, double> pointMap,
            double tolerance)
        {
            if (!GetEntityData(ent, out Point3d position, out string text))
                return false;

            if (!ParseElevation(text, out double z))
                return false;

            // Create XY key with tolerance
            Point3d key = new Point3d(
                Math.Round(position.X / tolerance) * tolerance,
                Math.Round(position.Y / tolerance) * tolerance,
                0
            );

            // Keep highest Z value for each XY location
            if (pointMap.TryGetValue(key, out double existingZ))
            {
                if (z > existingZ)
                    pointMap[key] = z;
            }
            else
            {
                pointMap.Add(key, z);
            }

            return true;
        }

        private bool GetEntityData(Entity ent, out Point3d position, out string text)
        {
            position = Point3d.Origin;
            text = string.Empty;

            switch (ent)
            {
                case DBText dbText:
                    position = dbText.Position;
                    text = CleanText(dbText.TextString);
                    return true;

                case MText mText:
                    position = mText.Location;
                    text = CleanText(mText.Contents);
                    return true;
            }
            return false;
        }

        private string CleanText(string input)
        {
            return Regex.Replace(input, @"\\[^;]+;|\{[^\}]+\}|[\r\n]+", " ")
                        .Replace("\\P", " ")
                        .Trim();
        }

        private bool ParseElevation(string text, out double z)
        {
            z = 0;
            Match match = Regex.Match(text, @"[-+]?\d+[.,]?\d*");
            return match.Success &&
                double.TryParse(match.Value.Replace(',', '.'),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out z);
        }

        private class Point2dComparer : IEqualityComparer<Point3d>
        {
            private readonly double _tolerance;

            public Point2dComparer(double tolerance) => _tolerance = tolerance;

            public bool Equals(Point3d a, Point3d b) =>
                Math.Abs(a.X - b.X) <= _tolerance &&
                Math.Abs(a.Y - b.Y) <= _tolerance;

            public int GetHashCode(Point3d p)
            {
                int xHash = Math.Round(p.X / _tolerance).GetHashCode();
                int yHash = Math.Round(p.Y / _tolerance).GetHashCode();
                return (xHash * 397) ^ yHash; // Manual hash code combination
            }
        }
    }
}