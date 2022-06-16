using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoomsPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddRoomsAndTags : IExternalCommand
    {
        // Команда автоматически создаёт в проекте набор помещений, помечая их марками.
        // Имя помещения, отобрадаемое в марке, содержит номер помещения в пределах этажа.
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            if (RoomHelper.IsRoomTagTypeSettingsFailed(doc))
            {
                return Result.Failed;
            }

            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();

            Transaction transaction = new Transaction(doc);
            transaction.Start("Добавление помещений");

            for (int levelNumber = 0; levelNumber < levels.Count; levelNumber++)
            {
                PlanTopology topology = doc.get_PlanTopology(levels[levelNumber]);
                PlanCircuitSet circuitSet = topology.Circuits;

                int roomNumber = 1;
                foreach (PlanCircuit circuit in circuitSet)
                {
                    if (!circuit.IsRoomLocated)
                    {
                        Room room = doc.Create.NewRoom(null, circuit);
                        room.Name = $"{levelNumber + 1}_{roomNumber}";
                        roomNumber++;
                    }
                }
            }

            transaction.Commit();

            return Result.Succeeded;
        }
    }

    [TransactionAttribute(TransactionMode.Manual)]
    public class RemoveAllRooms : IExternalCommand
    {
        // Команда удаляет из проекта все помещения
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            IList<ElementId> roomIds = new FilteredElementCollector(doc)
                .OfClass(typeof(SpatialElement))
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .ToElementIds() as IList<ElementId>;

            Transaction transaction = new Transaction(doc);
            transaction.Start("Удаление помещений");

            doc.Delete(roomIds);

            transaction.Commit();

            return Result.Succeeded;
        }
    }

    [TransactionAttribute(TransactionMode.Manual)]
    public class AddRoomTags : IExternalCommand
    {
        // Команда добавлет ко всем имеющимся в проекте помещениям марки, при этом изменяя имя помещения так,
        // чтобы оно содержало номер помещения в пределах этажа
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            if (RoomHelper.IsRoomTagTypeSettingsFailed(doc))
            {
                return Result.Failed;
            }

            List<Level> levels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();

            Transaction transaction = new Transaction(doc);
            transaction.Start("Добавление марок помещений");

            for (int levelNumber = 0; levelNumber < levels.Count; levelNumber++)
            {

                List<Room> rooms = new FilteredElementCollector(doc)
                    .OfClass(typeof(SpatialElement))
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .OfType<Room>()
                    .Where(r => r.LevelId == levels[levelNumber].Id)
                    .ToList();

                int roomNumber = 1;
                foreach (Room room in rooms)
                {
                    if (room.Area != 0)
                    {
                        room.Name = $"{levelNumber + 1}_{roomNumber}";
                        roomNumber++;
                        XYZ center = RoomHelper.GetRoomCenter(room);
                        UV centerUV = new UV(center.X, center.Y);
                        RoomTag roomTag = doc.Create.NewRoomTag(new LinkElementId(room.Id), centerUV, null);
                    }
                }
            }

            transaction.Commit();

            return Result.Succeeded;
        }
    }


    [TransactionAttribute(TransactionMode.Manual)]
    public class RemoveAllRoomTags : IExternalCommand
    {
        // Команда удаляет из проекта все марки помещений, оставляя сами помещения
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            IList<ElementId> roomTagIds = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .WhereElementIsNotElementType()
                .ToElementIds() as IList<ElementId>;

            Transaction transaction = new Transaction(doc);
            transaction.Start("Удаление марок помещений");

            doc.Delete(roomTagIds);

            transaction.Commit();

            return Result.Succeeded;
        }
    }

    // Класс вспомогательных методов
    public static class RoomHelper
    {
        public static XYZ GetRoomCenter(Room room)
        {
            XYZ boundCenter = GetElementCenter(room);
            LocationPoint locPoint = (LocationPoint)room.Location;
            XYZ roomCenter = new XYZ(boundCenter.X, boundCenter.Y, locPoint.Point.Z);
            return roomCenter;
        }

        public static XYZ GetElementCenter(Element elem)
        {
            BoundingBoxXYZ bounding = elem.get_BoundingBox(null);
            XYZ center = (bounding.Max + bounding.Min) * 0.5;
            return center;
        }

        // Метод проверяет, загружено ли в проект семейство Марки помещений и выполняет его настройку для отображения
        // только марки (имени помещения)
        public static bool IsRoomTagTypeSettingsFailed(Document doc)
        {
            FamilySymbol roomTagType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .OfType<FamilySymbol>()
                .FirstOrDefault();

            if (null == roomTagType)
            {
                TaskDialog.Show("Ошибка", "В проект не загружено семейство Марки помещений." +
                    "\nЗагрузите семейство и повторите операцию.");
                return true;
            }

            Transaction transaction = new Transaction(doc);
            transaction.Start("Установка параметров маркера помещения");

            roomTagType.LookupParameter("Показать номер помещения").Set(0);
            roomTagType.LookupParameter("Показать объем").Set(0);
            roomTagType.LookupParameter("Показать площадь").Set(0);

            transaction.Commit();

            return false;
        }
    }
}