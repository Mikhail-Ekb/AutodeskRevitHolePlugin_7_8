using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutodeskRevitHolePlugin_7_8
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document sysDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("R_4_78_Systems")).FirstOrDefault();
            if(sysDoc==null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл AutodeskRevitHolePlugin_7_8_Systems");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();
            if(familySymbol==null)
            {
                TaskDialog.Show("Ошибка", "Семейство Отверстия не найдено");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(sysDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();
            if (ducts == null)
            {
                TaskDialog.Show("Ошибка", "Воздуховоды не найдены");
                return Result.Cancelled;
            }

            List<Pipe> pipes = new FilteredElementCollector(sysDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();
            if (ducts == null)
            {
                TaskDialog.Show("Ошибка", "Трубы не найдены");
                return Result.Cancelled;
            }

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "3D вид не найден");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction ts0 = new Transaction(arDoc);
            ts0.Start("Расстановка отверстий");
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            ts0.Commit();


            Transaction ts = new Transaction(arDoc);
            ts.Start("Расстановка отверстий");
           
            foreach (Duct d in ducts)
            {
                Line curve_D = (d.Location as LocationCurve).Curve as Line;
                XYZ point_D = curve_D.GetEndPoint(0);
                XYZ direction_D = curve_D.Direction;

                List<ReferenceWithContext> intersections_D = referenceIntersector.Find(point_D, direction_D)
                    .Where(x => x.Proximity <= curve_D.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersections_D)
                {
                    double proximity_D = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole_D = point_D + (direction_D * proximity_D);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole_D, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width_D = hole.LookupParameter("Ширина");
                    Parameter hight_D = hole.LookupParameter("Высота");
                    width_D.Set(d.Diameter);
                    hight_D.Set(d.Diameter);
                }
            }

            foreach (Pipe p in pipes)
            {
                Line curve_P = (p.Location as LocationCurve).Curve as Line;
                XYZ point_P = curve_P.GetEndPoint(0);
                XYZ direction_P = curve_P.Direction;

                List<ReferenceWithContext> intersections_P = referenceIntersector.Find(point_P, direction_P)
                    .Where(x => x.Proximity <= curve_P.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();
                foreach (ReferenceWithContext refer in intersections_P)
                {
                    double proximity_P = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole_P = point_P + (direction_P * proximity_P);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole_P, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width_P = hole.LookupParameter("Ширина");
                    Parameter hight_P = hole.LookupParameter("Высота");
                    width_P.Set(p.Diameter);
                    hight_P.Set(p.Diameter);
                }
            }
            ts.Commit();
            return Result.Succeeded;
        }

       

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();
                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                    && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();
                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }
    }
}
