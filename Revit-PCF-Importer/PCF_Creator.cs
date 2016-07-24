using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using BuildingCoder;
using PCF_Functions;
using iv = PCF_Functions.InputVars;

namespace Revit_PCF_Importer
{
    public interface ICreateElements
    {
        Result SendElementsToCreation(ElementSymbol elementSymbol);
        }

    public interface IProcessElements
    {
        //Top level keywords
        Result ELEMENT_TYPE_NOT_IMPLEMENTED(ElementSymbol elementSymbol);
        Result PIPE(ElementSymbol elementSymbol);
        Result ELBOW(ElementSymbol elementSymbol);
}

    public class ProcessElements : IProcessElements
    {
        private Document doc = PCFImport.doc;
        public Result ELEMENT_TYPE_NOT_IMPLEMENTED(ElementSymbol elementSymbol)
        {
            return Result.Succeeded;
        }

        public Result PIPE(ElementSymbol elementSymbol)
        {
            try
            {
                //Choose pipe type. Hardcoded value until a configuring process is devised.
                ElementId pipeTypeId = new ElementId(3048519); //Hardcoded until configuring process is implemented

                //Collect levels and select one level
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ElementClassFilter levelFilter = new ElementClassFilter(typeof(Level));
                ElementId levelId = collector.WherePasses(levelFilter).FirstElementId();

                //Create pipe
                Pipe pipe = Pipe.Create(doc, elementSymbol.PipingSystemType.Id, pipeTypeId, levelId,
                    elementSymbol.EndPoint1.Xyz, elementSymbol.EndPoint2.Xyz);
                //Set pipe diameter
                Parameter parameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

                parameter.Set(elementSymbol.EndPoint1.Diameter);

                //Regenerate the docucment before trying to read a parameter that has been edited
                pipe.Document.Regenerate();
                //Store the reference for the created element in the ElementSymbol
                elementSymbol.CreatedElement = (Element) pipe;
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return Result.Failed;
            }
        }

        public Result ELBOW(ElementSymbol elementSymbol)
        {
            //Choose pipe type.Hardcoded value until a configuring process is devised.
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            Filter filter = new Filter("EN 10253-2 - Elbow: 3D", BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM); //Hardcoded until implements
            Element elbowSymbol = collector.WherePasses(filter.epf).FirstElement();
            try
            {
                //var query = (from ElementSymbol es in PCFImport.ExtractedElementCollection.Elements
                //             where string.Equals("PIPE", es.ElementType)
                //             select (MEPCurve)es.CreatedElement).ToList();

                //query.RemoveAll(item => item == null);

                //IList<Connector> allConnectors = new List<Connector>();

                //foreach (MEPCurve mepCurve in query)
                //{
                //    ConnectorSet conSet = mepCurve.ConnectorManager.Connectors;
                //    foreach (Connector c in conSet)
                //    {
                //        allConnectors.Add(c);
                //    }
                //}

                //XYZ p1 = elementSymbol.EndPoint1.Xyz; XYZ p2 = elementSymbol.EndPoint2.Xyz;

                //var c1 = (from Connector c in allConnectors where p1.IsAlmostEqualTo(c.Origin) select c).FirstOrDefault();
                //var c2 = (from Connector c in allConnectors where p2.IsAlmostEqualTo(c.Origin) select c).FirstOrDefault();

                //if (c1 != null || c2 != null)
                //{
                //    doc.Create.NewElbowFitting(c1, c2);
                //}


                if (elbowSymbol == null)
                {
                    Util.ErrorMsg("Family and Type for ELBOW at position " + elementSymbol.Position + " was not found.");
                    return Result.Failed;
                }

                Element element = doc.Create.NewFamilyInstance(elementSymbol.CentrePoint.Xyz, (FamilySymbol)elbowSymbol, StructuralType.NonStructural);

                //Begin geometric analysis to rotate the endpoints to actual locations
                //Get connectors from the placed family
                ConnectorSet cs = ((FamilyInstance)element).MEPModel.ConnectorManager.Connectors;

                Connector familyConnector1 = (from Connector c in cs where true select c).First();
                Connector familyConnector2 = (from Connector c in cs where true select c).Last();

                XYZ vA = familyConnector1.Origin - elementSymbol.CentrePoint.Xyz; //To define a vector: v = p2 - p1
                XYZ vC = elementSymbol.EndPoint1.Xyz - elementSymbol.CentrePoint.Xyz;

                XYZ normRotAxis = vC.CrossProduct(vA).Normalize();
                double dotProduct = vC.DotProduct(vA);
                double rotAngle = System.Math.Acos(dotProduct)*-1;
                var rotLine = Line.CreateUnbound(elementSymbol.CentrePoint.Xyz, normRotAxis);

                FamilyInstance fi = (FamilyInstance)element;

                fi.Location.Rotate(rotLine, rotAngle);

                //Store the reference of the created element in the symbol object.
                elementSymbol.CreatedElement = element;
                ;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return Result.Succeeded;
        }
    }
}
