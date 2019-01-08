using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Agw.Generic;
using Agw.Generic.ExtensionMethods;


namespace Agw.Coordinate
{
    public class XyLeader
    {
        [CommandMethod("AgwXyCoord")]
        public void AddCoordFlags()
        {
            var loop = true;
            while (loop)
            {
                loop = AddCoordLeader();
            }
        }

        private bool AddCoordLeader()
        {
            using (Active.Document.LockDocument())
            {
                var pnts = GetPointCollection();
                if (pnts == null || pnts.Count < 1) return false;
                pnts.Reverse();

                var mLeaderStyleId = createMleaderStyleAndSetCurrent();
                if (mLeaderStyleId == ObjectId.Null) return false;

                return DrawLeader(mLeaderStyleId, pnts);
            }

        }

        private bool DrawLeader(ObjectId mLeaderStyleId, IList<Point3d> pnt)
        {
            var leader = new MLeader()
            {
                MLeaderStyle = mLeaderStyleId,
                Layer = "0",
            };

            using (Active.Document.LockDocument())
            using (Transaction tr = Active.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    var idx = leader.AddLeaderLine(pnt[0]);
                    var direction = 1.0;
                    if ((pnt[0] - pnt[1]).DotProduct(Vector3d.XAxis) < 0)
                    {
                        direction = -1;
                    }

                    leader.SetDogleg(idx, Vector3d.XAxis.MultiplyBy(direction));

                    var text = "{X= " + Decimal.Round((decimal)pnt[0].X, 2) + "}" +
                               "\\P{Y= " + Decimal.Round((decimal)pnt[0].Y, 2) + "}";

                    pnt.RemoveAt(0);

                    foreach (Point3d pt in pnt)
                    {
                        leader.AddFirstVertex(idx, pt);
                    }

                    MText mText = new MText();
                    mText.SetDatabaseDefaults();

                    mText.Annotative = AnnotativeStates.True;
                    mText.Contents = text;
                    mText.Attachment = (direction < 0) ? AttachmentPoint.MiddleRight : AttachmentPoint.MiddleLeft;
                    leader.MText = mText;
                    leader.TextHeight = 1.75;
                    leader.EnableFrameText = true;

                    Active.Database.AddEntities(leader);
                    tr.Commit();
                    return true;
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e);
                    return false;
                }
            }
        }

        private ObjectId createMleaderStyleAndSetCurrent()
        {
            var styleName = "COORDS";
            var mLeaderStyle = ObjectId.Null;

            using (Active.Document.LockDocument())
            using (Transaction tr = Active.Database.TransactionManager.StartTransaction())
            {
                var mlStyles = tr.GetObject(Active.Database.MLeaderStyleDictionaryId, OpenMode.ForRead) as DBDictionary;

                if (!mlStyles.Contains(styleName))
                {
                    var newStyle = new MLeaderStyle
                    {
                        Annotative = AnnotativeStates.True,
                        LeaderLineType = LeaderType.StraightLeader,
                        LandingGap = 2,
                        EnableDogleg = true,
                        DoglegLength = 2,
                        ArrowSize = 2,
                        BreakSize = 2,
                        ContentType = ContentType.MTextContent
                    };

                    mLeaderStyle = newStyle.PostMLeaderStyleToDb(Active.Database, styleName);
                    tr.AddNewlyCreatedDBObject(newStyle, true);
                }
                else
                {
                    mLeaderStyle = mlStyles.GetAt(styleName);
                }
                Active.Database.MLeaderstyle = mLeaderStyle;
                tr.Commit();
            }
            return mLeaderStyle;
        }

        private List<Point3d> GetPointCollection()
        {
            var returnvalue = new List<Point3d>();
            var pOpt = new PromptPointOptions(Environment.NewLine + "Select Point") { AllowNone = true };
            var result = Active.Editor.GetPoint(pOpt);
            if (result.Status != PromptStatus.OK || result.Status == PromptStatus.None) return null;

            returnvalue.Add(result.Value);

            pOpt.Message = Environment.NewLine + "Select label pos:";
            pOpt.UseBasePoint = true;
            pOpt.BasePoint = result.Value;
            var result1 = Active.Editor.GetPoint(pOpt);
            if (result1.Status != PromptStatus.OK || result.Status == PromptStatus.None) return null;

            returnvalue.Add(result1.Value);

            return returnvalue;
        }
    }
}