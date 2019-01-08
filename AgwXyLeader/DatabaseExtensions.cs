using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace Agw.Generic.ExtensionMethods
{
    public static class DatabaseExtensions
    {
        /// <summary>
        /// Loop over all Entities of Type T in Modelspace
        /// Example  Active.Database.ForEach<Line>(lijn =>{if (lijn.Length < 100.0){ lijn.UpgradeOpen(); lijn.ColorIndex = 1; } }  );
        ///  </summary>
        /// <typeparam name="T"> all object derived of Entity</typeparam>
        /// <param name="database"></param>
        /// <param name="action"> Method with parameter entity</param>
        public static void ForEach<T>(this Database database, Action<T> action) where T : Entity
        {

            using (var tr = database.TransactionManager.StartTransaction())
            {
                // Get the block table for the current database
                var blockTable = (BlockTable)tr.GetObject(database.BlockTableId, OpenMode.ForRead);

                // Get the model space block table record
                var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // get theClassType
                RXClass theClass = RXObject.GetClass(typeof(T));

                // Loop through the entities in model space
                foreach (ObjectId objectId in modelSpace)
                {
                    // Look for entities of the correct type
                    if (objectId.ObjectClass.IsDerivedFrom(theClass))
                    {
                        var entity =
                            (T)tr.GetObject(
                                objectId, OpenMode.ForWrite);

                        action(entity);
                    }
                }
                tr.Commit();
            }
        }

        public static void ForEachInAllBlockTableRecords<T>(this Database database, Action<T> action) where T : Entity
        {

            using (var tr = database.TransactionManager.StartTransaction())
            {
                // Get the block table for the current database
                var blockTable = (BlockTable)tr.GetObject(database.BlockTableId, OpenMode.ForRead);

                foreach (ObjectId id in blockTable)
                {
                    // Get the model space block table record
                    var blockTableRecord = (BlockTableRecord)tr.GetObject(id, OpenMode.ForRead);
                    // get theClassType
                    RXClass theClass = RXObject.GetClass(typeof(T));

                    // Loop through the entities in model space
                    foreach (ObjectId objectId in blockTableRecord)
                    {
                        // Look for entities of the correct type
                        if (objectId.ObjectClass.IsDerivedFrom(theClass))
                        {
                            var entity =
                                (T)tr.GetObject(
                                    objectId, OpenMode.ForWrite);

                            action(entity);
                        }
                    }

                }
                tr.Commit();
            }
        }

        public static ObjectId AddEntities(this Database database, Entity entity)
        {
            // Start a transaction
            using (DocumentLock lk = Active.Document.LockDocument())
            using (Transaction tr = database.TransactionManager.StartTransaction())
            {
                // Get the block table for the current database
                var blockTable = (BlockTable)tr.GetObject(database.BlockTableId, OpenMode.ForRead);

                // Get the model space block table record
                var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Add the new object to the block table record and the transaction
                ObjectId returnvalue = modelSpace.AppendEntity(entity);
                tr.AddNewlyCreatedDBObject(entity, true);

                // Save the new object to the database
                tr.Commit();

                return returnvalue;
            }
        }

        public static List<ObjectId> AddEntities(this Database database, List<Entity> entities)
        {
            List<ObjectId> returnvalue = new List<ObjectId>();
            // Start a transaction
            using (DocumentLock lk = Active.Document.LockDocument())
            using (Transaction tr = database.TransactionManager.StartTransaction())
            {
                // Get the block table for the current database
                var blockTable = (BlockTable)tr.GetObject(database.BlockTableId, OpenMode.ForRead);

                // Get the model space block table record
                var modelSpace = (BlockTableRecord)tr.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Add the new object to the block table record and the transaction
                foreach (Entity entity in entities)
                {
                    returnvalue.Add(modelSpace.AppendEntity(entity));
                    tr.AddNewlyCreatedDBObject(entity, true);
                }

                // Save the new objects to the database
                tr.Commit();

                //Return the objectId's for further use
                return returnvalue;
            }
        }
    }
}
