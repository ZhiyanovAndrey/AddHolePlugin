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

namespace AddHolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document arDoc = uidoc.Document;

            //поиск фаила документа содержащего слово ОВ. Возьмем первогый в списке или по умолчанию
            Document ovDoc = arDoc.Application.Documents
                .OfType<Document>()
                .Where(x => x.Title.Contains("ОВ"))
                .FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }
            //поиск созданного семейства. Возьмем первогый в списке или по умолчанию
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстие"))
                .FirstOrDefault();

            //если не найдено то сообщим и выйдем из плагина
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстие\"");
                return Result.Cancelled;
            }
            //фильтрация воздуховодов
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();
            //фильтрация труб
            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                 .OfClass(typeof(Pipe))
                 .OfType<Pipe>()
                 .ToList();

            //что бы применить фильтр ReferenceIntersector, работает только на 3D виде, найдем 3D вид

            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate) //условие что свойство IsTemplate не установленно
                .FirstOrDefault();
            //если не найдено то сообщим и выйдем из плагина
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)),
                FindReferenceTarget.Element, view3D);

            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");

            GetHoleDuct(arDoc, familySymbol, ducts, referenceIntersector);
            GetHolePipe(arDoc, familySymbol, pipes, referenceIntersector);

            transaction.Commit();

            return Result.Succeeded;
        }

        private void GetHolePipe(Document arDoc, FamilySymbol familySymbol, List<Pipe> pipes, ReferenceIntersector referenceIntersector)
        {
            foreach (Pipe pipe in pipes)
            {                                                               //ElementClassFilter и OfClass делают одно и то же
                Line curve = (pipe.Location as LocationCurve).Curve as Line;  //у воздуховода определяем Location обращаемся к кривой результат заносим в переменную
                                                                              //так как у прямых Line есть св-во Direction
                XYZ point = curve.GetEndPoint(0); //найдем точку
                XYZ direction = curve.Direction;

                //найдем пересечения линий со стеной, ограничим колекцию элементами с длиой меньше воздуховода
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity; //расстояние до стены
                    Reference reference = refer.GetReference(); //узнаем id элемента
                    //зная id можно получить в докум стену на которую идет ссылка
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; //получим стену
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //получим уровень стены

                    XYZ pointHole = point + (direction * proximity); //точка начала+(направление луча*расстояние до стены direction имеет тип XYZ)

                    //экземпляр семейства добавится с определенными нами размерами
                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                    //заменим параметры
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(pipe.Diameter);
                    height.Set(pipe.Diameter);
                }
            }
        }

        private static void GetHoleDuct(Document arDoc, FamilySymbol familySymbol, List<Duct> ducts, ReferenceIntersector referenceIntersector)
        {
            foreach (Duct duct in ducts)
            {                                                               //ElementClassFilter и OfClass делают одно и то же
                Line curve = (duct.Location as LocationCurve).Curve as Line;  //у воздуховода определяем Location обращаемся к кривой результат заносим в переменную
                                                                              //так как у прямых Line есть св-во Direction
                XYZ point = curve.GetEndPoint(0); //найдем точку
                XYZ direction = curve.Direction;

                //найдем пересечения линий со стеной, ограничим колекцию элементами с длиой меньше воздуховода
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity; //расстояние до стены
                    Reference reference = refer.GetReference(); //узнаем id элемента
                    //зная id можно получить в докум стену на которую идет ссылка
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; //получим стену
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //получим уровень стены

                    XYZ pointHole = point + (direction * proximity); //точка начала+(направление луча*расстояние до стены direction имеет тип XYZ)

                    //экземпляр семейства добавится с определенными нами размерами
                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);

                    //заменим параметры
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(duct.Diameter);
                    height.Set(duct.Diameter);
                }
            }
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            //определяет имеют ли два объекта точки на одном элементе (одной стене) то true
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

            //метод возвращает HashCode обьекта
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
