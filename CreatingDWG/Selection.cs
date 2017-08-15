using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace CreatingDWG
{
    class Selection
    {
        /// <summary>
        /// Select all line references contained in the input layer
        /// </summary>
        /// <param name="doc">the current autocad document</param>
        /// <returns></returns>
        public static ObjectIdCollection SelectEntities(Document doc, string LayerName)
        {
            Editor ed = doc.Editor;
            ObjectIdCollection collection = new ObjectIdCollection();

            // Build a conditional filter list so that only entities with the specified properties are selected

            TypedValue[] tvs = new TypedValue[] {
          new TypedValue((int)DxfCode.Operator,"<or"),
          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"POLYLINE"),
          new TypedValue((int)DxfCode.Operator, "and>"),

                    new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"LINE"),
          new TypedValue((int)DxfCode.Operator, "and>"),

          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"DIMENSION"),
          new TypedValue((int)DxfCode.Operator, "and>"),

          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"CIRCLE"),
          new TypedValue((int)DxfCode.Operator, "and>"),

          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"MTEXT"),
          new TypedValue((int)DxfCode.Operator, "and>"),
          new TypedValue((int)DxfCode.Operator,"or>")
        };

            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            List<string> log = new List<string>();

            if (psr.Status == PromptStatus.OK)
            {
                foreach (SelectedObject obj in psr.Value)
                {
                    collection.Add(obj.ObjectId);
                }
            }

            //ed.WriteMessage("\nFound {0} entit{1}.", psr.Value.Count,(psr.Value.Count == 1 ? "y" : "ies"));

            return collection;
        }

        public static ObjectIdCollection SelectCyanLines(Document doc, string LayerName)
        {
            Editor ed = doc.Editor;
            ObjectIdCollection collection = new ObjectIdCollection();

            // Build a conditional filter list so that only entities with the specified properties are selected
            TypedValue[] tvs = new TypedValue[] {
                new TypedValue((int)DxfCode.Operator,"<or"),

          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"LINE"),
          new TypedValue((int)DxfCode.Color,4),
          new TypedValue((int)DxfCode.Operator, "and>"),

                    new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"LINE"),
          new TypedValue((int)DxfCode.Color,2),
          new TypedValue((int)DxfCode.Operator, "and>"),

          new TypedValue((int)DxfCode.Operator,"or>")
        };

            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            List<string> log = new List<string>();

            if (psr.Status == PromptStatus.OK)
            {
                foreach (SelectedObject obj in psr.Value)
                {
                    collection.Add(obj.ObjectId);
                }
            }

            return collection;
        }

        public static ObjectIdCollection SelectGreenLines(Document doc, string LayerName)
        {
            Editor ed = doc.Editor;
            ObjectIdCollection collection = new ObjectIdCollection();

            // Build a conditional filter list so that only entities with the specified properties are selected
            TypedValue[] tvs = new TypedValue[] {
                new TypedValue((int)DxfCode.Operator,"<or"),

          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"LINE"),
          new TypedValue((int)DxfCode.Color,3),
          new TypedValue((int)DxfCode.Operator, "and>"),

                    new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"LINE"),
          new TypedValue((int)DxfCode.Color,2),
          new TypedValue((int)DxfCode.Operator, "and>"),

          new TypedValue((int)DxfCode.Operator,"or>")
        };

            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            List<string> log = new List<string>();

            if (psr.Status == PromptStatus.OK)
            {
                foreach (SelectedObject obj in psr.Value)
                {
                    collection.Add(obj.ObjectId);
                }
            }

            return collection;
        }

        public static ObjectIdCollection CheckingCollection(Document doc, string LayerName)
        {
            Editor ed = doc.Editor;
            ObjectIdCollection collection = new ObjectIdCollection();

            // Build a conditional filter list so that only entities with the specified properties are selected
            TypedValue[] tvs = new TypedValue[] {
          new TypedValue((int)DxfCode.Operator,"<and"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Start,"LINE"),
          new TypedValue((int)DxfCode.Color,3),
          new TypedValue((int)DxfCode.Operator, "and>"),
        };

            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            List<string> log = new List<string>();

            if (psr.Status == PromptStatus.OK)
            {
                foreach (SelectedObject obj in psr.Value)
                {
                    collection.Add(obj.ObjectId);
                }
            }

            return collection;
        }

        public static ObjectIdCollection SelectEntitiesWithinLayer(Document doc, string LayerName)
        {
            Editor ed = doc.Editor;
            ObjectIdCollection collection = new ObjectIdCollection();

            // Build a conditional filter list so that only entities with the specified properties are selected

            TypedValue[] tvs = new TypedValue[] {
          new TypedValue((int)DxfCode.Operator,"<or"),
          new TypedValue((int)DxfCode.LayerName,LayerName),
          new TypedValue((int)DxfCode.Operator,"or>")
        };

            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            List<string> log = new List<string>();

            if (psr.Status == PromptStatus.OK)
            {
                foreach (SelectedObject obj in psr.Value)
                {
                    collection.Add(obj.ObjectId);
                }
            }

            return collection;
        }

        public static ObjectIdCollection SelectNameText(Document doc, string LayerName)
        {
            Editor ed = doc.Editor;
            ObjectIdCollection collection = new ObjectIdCollection();

            // Build a conditional filter list so that only entities with the specified properties are selected

            TypedValue[] tvs = new TypedValue[] {
          new TypedValue((int)DxfCode.Operator,"<or"),
          new TypedValue((int)DxfCode.LayerName,"07-Name"),
          new TypedValue((int)DxfCode.Operator,"or>")
        };

            SelectionFilter sf = new SelectionFilter(tvs);
            PromptSelectionResult psr = ed.SelectAll(sf);

            List<string> log = new List<string>();

            if (psr.Status == PromptStatus.OK)
            {
                foreach (SelectedObject obj in psr.Value)
                {
                    collection.Add(obj.ObjectId);
                }
            }

            return collection;
        }


    }
}
