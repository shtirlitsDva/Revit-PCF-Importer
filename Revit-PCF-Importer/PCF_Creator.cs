using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.AccessControl;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using BuildingCoder;
using PCF_Functions;
using iv = PCF_Functions.InputVars;
using ch = PCF_Functions.CreatorHelper;
using ex = BuildingCoder.Extensions;
using pif = Revit_PCF_Importer.PCF_Importer_form;

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
        Result TEE(ElementSymbol elementSymbol);
        Result CAP(ElementSymbol elementSymbol);
        Result FLANGE(ElementSymbol elementSymbol);
        Result FLANGE_BLIND(ElementSymbol elementSymbol);
        Result REDUCER_CONCENTRIC(ElementSymbol elementSymbol);
        Result OLET(ElementSymbol elementSymbol);
        Result VALVE(ElementSymbol elementSymbol);
    }

    public class ProcessElements : IProcessElements
    {
        //private PCFImport.document PCFImport.doc = PCFImport.doc;
        public Result ELEMENT_TYPE_NOT_IMPLEMENTED(ElementSymbol elementSymbol)
        {
            return Result.Succeeded;
        }

        public Result PIPE(ElementSymbol elementSymbol)
        {
            try
            {
                //Choose pipe type.
                ElementId pipeTypeId = elementSymbol.PipeType.Id;

                //Collect levels and select one level
                FilteredElementCollector collector = new FilteredElementCollector(PCFImport.doc);
                ElementClassFilter levelFilter = new ElementClassFilter(typeof(Level));
                ElementId levelId = collector.WherePasses(levelFilter).FirstElementId();

                //Test if pipe shorter than 2 mm, if true abort the creation
                if (Util.MinPipeLength > Util.Distance(elementSymbol.EndPoint1.Xyz, elementSymbol.EndPoint2.Xyz)) return Result.Failed;

                //Create pipe
                Pipe pipe = Pipe.Create(PCFImport.doc, elementSymbol.PipingSystemType.Id, pipeTypeId, levelId,
                    elementSymbol.EndPoint1.Xyz, elementSymbol.EndPoint2.Xyz);
                //Set pipe diameter
                Parameter parameter = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM);

                parameter.Set(elementSymbol.EndPoint1.Diameter);

                //Store the reference for the created element in the ElementSymbol
                elementSymbol.CreatedElement = (Element)pipe;
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
            try
            {
                //Choose pipe type.Hardcoded value until a configuring process is devised.
                //FilteredElementCollector collector = new FilteredElementCollector(PCFImport.doc);
                //Filter filter = new Filter("EN 10253-2 - Elbow: 3D", BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM); //Hardcoded until implements
                //Applying a fast filter first (OfCategory) to reduce load on slow filter (WherePasses).
                //FamilySymbol elbowSymbol = collector.OfCategory(BuiltInCategory.OST_PipeFitting).WherePasses(filter.epf).Cast<FamilySymbol>().FirstOrDefault();
                //if (elbowSymbol == null)
                //{
                //    Util.ErrorMsg("Family and Type for ELBOW at position " + elementSymbol.Position + " was not found.");
                //    return Result.Failed;
                //}

                #region Placing by face reference

                ////Build a direct shape with TessellatedShapeBuilder
                //List<XYZ> args = new List<XYZ>(3);

                //TessellatedShapeBuilder builder = new TessellatedShapeBuilder();
                ////http://thebuildingcoder.typepad.com/blog/2014/05/directshape-performance-and-minimum-size.html
                //builder.OpenConnectedFaceSet(false);
                //args.Add(elementSymbol.CentrePoint.Xyz);
                //args.Add(elementSymbol.EndPoint1.Xyz);
                //args.Add(elementSymbol.EndPoint2.Xyz);
                //builder.AddFace(new TessellatedFace(args,ElementId.InvalidElementId));
                //builder.CloseConnectedFaceSet();
                //builder.Build();
                //TessellatedShapeBuilderResult result = builder.GetBuildResult();

                //var resultList = result.GetGeometricalObjects();

                //var solidShape = resultList[0] as Solid;
                //Face face = solidShape.Faces.get_Item(0);

                //DirectShape ds = DirectShape.CreateElement(PCFImport.doc, new ElementId(BuiltInCategory.OST_GenericModel));
                //ds.ApplicationId = "Application id";
                //ds.ApplicationDataId = "Geometry object id";
                //ds.Name = "Elbow " + elementSymbol.Position;
                //DirectShapeOptions dso = ds.GetOptions();
                //dso.ReferencingOption = DirectShapeReferencingOption.Referenceable;
                //ds.SetOptions(dso);
                //ds.SetShape(resultList);
                //Options options = new Options();
                //options.ComputeReferences = true;
                //PCFImport.doc.Regenerate();
                //FilteredElementCollector collectorDs = new FilteredElementCollector(PCFImport.doc);
                //collectorDs.OfClass(typeof (DirectShape));
                //var query = from Element e in collectorDs
                //    where string.Equals(e.Name, "Elbow " + elementSymbol.Position)
                //    select e;
                //Element elem = query.FirstOrDefault();
                //var geometryElement = elem.get_Geometry(options);
                //;
                //Face face = null;
                //foreach (GeometryObject geometry in geometryElement)
                //{
                //    Solid instance = geometry as Solid;
                //    if (null == instance || 0 == instance.Faces.Size || 0 == instance.Edges.Size) { continue; }
                //    // Get the faces
                //    foreach (Face f in instance.Faces)
                //    {
                //        face = f;
                //    }
                //}
                //;
                //if (face == null)
                //{
                //    Util.ErrorMsg("No valid face detected to place the fitting for element at position "+elementSymbol.Position);
                //    return Result.Failed;
                //}
                //var faceRef = face.Reference;
                //;
                ////Finally place the elbow
                ////Direction -- third parameter to the create method
                //XYZ vC = elementSymbol.EndPoint1.Xyz - elementSymbol.CentrePoint.Xyz;
                //Element element = PCFImport.doc.Create.NewFamilyInstance(faceRef, elementSymbol.CentrePoint.Xyz, vC,
                //    (FamilySymbol) elbowSymbol);

                #endregion

                #region Create by NewElbowFitting
                //Get all pipe connectors
                HashSet<Connector> allPipeConnectors = ch.GetAllPipeConnectors();

                //Get the actual endpoints of the elbow
                XYZ p1 = elementSymbol.EndPoint1.Xyz; XYZ p2 = elementSymbol.EndPoint2.Xyz;
                //Determine the corresponding pipe connectors
                var c1 = (from Connector c in allPipeConnectors where Util.IsEqual(p1, c.Origin) select c).FirstOrDefault();
                var c2 = (from Connector c in allPipeConnectors where Util.IsEqual(p2, c.Origin) select c).FirstOrDefault();

                //Handle the missing connectors by creating dummy pipes

                Pipe pipe1 = null; Pipe pipe2 = null;

                if (c1 == null)
                {
                    pipe1 = ch.CreateDummyPipe(p1, elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint1, elementSymbol);
                    c1 = ch.MatchConnector(p1, pipe1);
                }

                if (c2 == null)
                {
                    pipe2 = ch.CreateDummyPipe(p2, elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint2, elementSymbol);
                    c2 = ch.MatchConnector(p2, pipe2);
                }

                if (c1 != null && c2 != null)
                {
                    Element element = PCFImport.doc.Create.NewElbowFitting(c1, c2);
                    if (pipe1 != null) PCFImport.doc.Delete(pipe1.Id);
                    if (pipe2 != null) PCFImport.doc.Delete(pipe2.Id);
                    elementSymbol.CreatedElement = element;
                    PCFImport.doc.Regenerate();
                    return Result.Succeeded;
                }
                //If this point is reached, something has failed
                return Result.Failed;
                #endregion

                #region Create by Directly Placing families

                //Element element = PCFImport.doc.Create.NewFamilyInstance(elementSymbol.CentrePoint.Xyz, (FamilySymbol)elbowSymbol, StructuralType.NonStructural);

                //FamilyInstance elbow = (FamilyInstance)element;

                //double diameter = elementSymbol.EndPoint1.Diameter;

                //elbow.LookupParameter("Nominal Diameter").Set(diameter); //Implement a procedure to select the parameter by name supplied by user

                ////Begin geometric analysis to rotate the endpoints to actual locations
                ////Get connectors from the placed family
                //ConnectorSet cs = ((FamilyInstance)element).MEPModel.ConnectorManager.Connectors;

                //Connector familyConnector1 = (from Connector c in cs where true select c).First();
                //Connector familyConnector2 = (from Connector c in cs where true select c).Last();

                //XYZ vA = familyConnector1.Origin - elementSymbol.CentrePoint.Xyz; //To define a vector: v = p2 - p1
                //XYZ vC = elementSymbol.EndPoint1.Xyz - elementSymbol.CentrePoint.Xyz;

                //XYZ vB = familyConnector2.Origin - elementSymbol.CentrePoint.Xyz; //To define a vector: v = p2 - p1
                //XYZ vD = elementSymbol.EndPoint2.Xyz - elementSymbol.CentrePoint.Xyz;

                //XYZ normRotAxis = vC.CrossProduct(vA).Normalize();

                //#region Fun with model lines

                //Filter filtermarker = new Filter("Marker: Marker", BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM); //Hardcoded until implements
                //FamilySymbol markerSymbol = new FilteredElementCollector(PCFImport.doc).WherePasses(filtermarker.epf).Cast<FamilySymbol>().FirstOrDefault();

                //XYZ A = vC.CrossProduct(vA);
                //XYZ B = vD.CrossProduct(vB);

                //Element marker = PCFImport.doc.Create.NewFamilyInstance(elementSymbol.CentrePoint.Xyz.Add(A), markerSymbol, StructuralType.NonStructural);
                ////Helper.PlaceAdaptiveMarkerLine("Red", elementSymbol.CentrePoint.Xyz, elementSymbol.CentrePoint.Xyz.Add(A));
                ////Helper.PlaceAdaptiveMarkerLine("Green", elementSymbol.CentrePoint.Xyz, elementSymbol.CentrePoint.Xyz.Add(B));


                //#endregion
                //Line rotLine = Line.CreateBound(elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint1.Xyz);
                //double rotAngle = Math.PI;

                ////double dotProduct = vC.DotProduct(vA);
                ////double rotAngle = System.Math.Acos(dotProduct);
                ////var rotLine = Line.CreateUnbound(elementSymbol.CentrePoint.Xyz, normRotAxis); //Rotation line must be BOUND!!!!

                //////Test rotation
                ////Transform trf = Transform.CreateRotationAtPoint(normRotAxis, rotAngle, elementSymbol.CentrePoint.Xyz);
                ////XYZ testRotation = trf.OfVector(vA).Normalize();

                ////if ((vC.DotProduct(testRotation) > 0.00001) == false) rotAngle = -rotAngle;


                //elbow.Location.Rotate(rotLine, rotAngle);

                ////Store the reference of the created element in the symbol object.
                //elementSymbol.CreatedElement = element;
                //;

                #endregion
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new Exception(e.Message);
            }
        }

        public Result TEE(ElementSymbol elementSymbol)
        {
            try
            {
                //Determine if the tee is reducing. Move this to PARSER?
                if (!elementSymbol.Branch1Point.Diameter.Equals(elementSymbol.EndPoint1.Diameter)) elementSymbol.IsReducing = true;

                #region ByPlacingFamilyInstance
                //Filter filter;
                //if (!elementSymbol.IsReducing) filter = new Filter("EN 10253-2 - Tee: Tee Type B", BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                //else filter = new Filter("EN 10253-2 - Reducing Tee: Red Tee Type B", BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM);
                //FamilySymbol teeSymbol = collector.WherePasses(filter.epf).Cast<FamilySymbol>().FirstOrDefault();

                //if (teeSymbol == null)
                //{
                //    Util.ErrorMsg("Family and Type for TEE at position " + elementSymbol.Position + " was not found.");
                //    return Result.Failed;
                //}

                //Element element = PCFImport.doc.Create.NewFamilyInstance(elementSymbol.CentrePoint.Xyz, teeSymbol, StructuralType.NonStructural);

                //FamilyInstance tee = (FamilyInstance)element;

                //double mainDiameter = elementSymbol.EndPoint1.Diameter;
                //double branchDiameter = elementSymbol.Branch1Point.Diameter;

                //tee.LookupParameter("Nominal Diameter 1").Set(mainDiameter); //Implement a procedure to select the parameter by name supplied by user
                //if (elementSymbol.IsReducing) tee.LookupParameter("Nominal Diameter 3").Set(branchDiameter);
                #endregion

                //Get all pipe connectors
                HashSet<Connector> allPipeConnectors = ch.GetAllPipeConnectors();

                //Get the actual endpoints of the elbow
                XYZ p1 = elementSymbol.EndPoint1.Xyz; XYZ p2 = elementSymbol.EndPoint2.Xyz; XYZ p3 = elementSymbol.Branch1Point.Xyz;
                //Determine the corresponding pipe connectors
                var c1 = (from Connector c in allPipeConnectors where Util.IsEqual(p1, c.Origin) select c).FirstOrDefault();
                var c2 = (from Connector c in allPipeConnectors where Util.IsEqual(p2, c.Origin) select c).FirstOrDefault();
                var c3 = (from Connector c in allPipeConnectors where Util.IsEqual(p3, c.Origin) select c).FirstOrDefault();

                //Handle the missing connectors by creating dummy pipes

                Pipe pipe1 = null; Pipe pipe2 = null; Pipe pipe3 = null;

                if (c1 == null)
                {
                    pipe1 = ch.CreateDummyPipe(p1, elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint1, elementSymbol);
                    c1 = ch.MatchConnector(p1, pipe1);
                }

                if (c2 == null)
                {
                    pipe2 = ch.CreateDummyPipe(p2, elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint2, elementSymbol);
                    c2 = ch.MatchConnector(p2, pipe2);
                }

                if (c3 == null)
                {
                    pipe3 = ch.CreateDummyPipe(p3, elementSymbol.CentrePoint.Xyz, elementSymbol.Branch1Point, elementSymbol);
                    c3 = ch.MatchConnector(p3, pipe3);
                }

                if (c1 != null && c2 != null && c3 != null)
                {
                    Element element = PCFImport.doc.Create.NewTeeFitting(c1, c2, c3);
                    if (pipe1 != null) PCFImport.doc.Delete(pipe1.Id);
                    if (pipe2 != null) PCFImport.doc.Delete(pipe2.Id);
                    if (pipe3 != null) PCFImport.doc.Delete(pipe3.Id);
                    elementSymbol.CreatedElement = element;
                    PCFImport.doc.Regenerate();
                    return Result.Succeeded;
                }
                //If this point is reached, something has failed
                return Result.Failed;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw new Exception(e.Message);
            }
        }

        public Result CAP(ElementSymbol elementSymbol)
        {
            try
            {
                //The CAP has two end-points and no way to know which one of them to use to place the family
                //Sooo... we match both end points to existing connectors in project to determine which one of them to use

                FilteredElementCollector allElementsWithConnectors =
                    ch.GetElementsWithConnectors(PCFImport.doc);

                XYZ firstMatch = null;
                XYZ secondMatch = null;

                firstMatch = (
                    from elem in allElementsWithConnectors
                        //Get all elements with connectors in PCFImport.document in a collector
                    select ch.GetConnectorSet(elem)
                    //Retrieve the connector set of each element in collector
                    into connectorSet //Pass it on
                    from Connector c in connectorSet //Declare that we are looking at the connectors
                    where Util.IsEqual(elementSymbol.EndPoint1.Xyz, c.Origin)
                    //Compare the location from the symbol to each connector in the PCFImport.document
                    select c.Origin).FirstOrDefault(); //Break on first match

                secondMatch = (
                    from elem in allElementsWithConnectors
                        //Get all elements with connectors in PCFImport.document in a collector
                    select ch.GetConnectorSet(elem)
                    //Retrieve the connector set of each element in collector
                    into connectorSet //Pass it on
                    from Connector c in connectorSet //Declare that we are looking at the connectors
                    where Util.IsEqual(elementSymbol.EndPoint2.Xyz, c.Origin)
                    //Compare the location from the symbol to each connector in the PCFImport.document
                    select c.Origin).FirstOrDefault(); //Break on first match
                //If no matching location is found -- fail the operation
                if (firstMatch == null && secondMatch == null)
                {
                    Util.ErrorMsg("Placement location for CAP at position " + elementSymbol.Position +
                                  " could not be determined.");
                    return Result.Failed;
                }
                //Select the correct location for placement assuming always only one correct match
                XYZ placementLocation = null;
                PointInSpace placementEnd = null;
                XYZ otherLocation = null;
                PointInSpace otherEnd = null;

                if (firstMatch != null)
                {
                    placementLocation = firstMatch;
                    placementEnd = elementSymbol.EndPoint1;
                    otherLocation = elementSymbol.EndPoint2.Xyz;
                    otherEnd = elementSymbol.EndPoint2;
                }
                else
                {
                    placementLocation = secondMatch;
                    placementEnd = elementSymbol.EndPoint2;
                    otherLocation = elementSymbol.EndPoint1.Xyz;
                    otherEnd = elementSymbol.EndPoint1;
                }


                //Place the instance
                Element cap = PCFImport.doc.Create.NewFamilyInstance(placementLocation, elementSymbol.FamilySymbol,
                    StructuralType.NonStructural);

                ConnectorSet conSet = ch.GetConnectorSet(cap);
                //The CAP should only have one connector
                Connector c1 = (from Connector c in conSet where true select c).FirstOrDefault();

                Pipe pipe1 = null;

                //Get all pipe connectors
                HashSet<Connector> allPipeConnectors = ch.GetAllPipeConnectors();

                //Determine the corresponding pipe connectors
                Connector c2 =
                    (from Connector c in allPipeConnectors where Util.IsEqual(placementLocation, c.Origin) select c)
                        .FirstOrDefault();

                if (c2 != null) pipe1 = c2.Owner as Pipe;

                ////Find the other connector (again...)
                //Connector c2 = (
                //    from elem in allElementsWithConnectors //Get all elements with connectors in PCFImport.document in a collector
                //    select CreatorHelper.GetConnectorSet(elem) //Retrieve the connector set of each element in collector
                //    into connectorSet //Pass it on
                //    from Connector c in connectorSet //Declare that we are looking at the connectors
                //    where Util.IsEqual(c1.Origin, c.Origin) //Compare the connector from the cap to each connector in the PCFImport.document
                //    select c).FirstOrDefault(); //Break on first match

                //Create a dummy pipe to attach the cap to
                if (c2 == null)
                {
                    pipe1 = ch.CreateDummyPipe(placementLocation, otherLocation, placementEnd, elementSymbol);
                    c2 = ch.MatchConnector(placementLocation, pipe1);
                    elementSymbol.DummyToDelete = pipe1;
                }

                #region Geometric manipulation
                //http://thebuildingcoder.typepad.com/blog/2012/05/create-a-pipe-cap.html
                Connector capConnector = c1;
                Connector start = c2;
                //Select the OTHER connector
                MEPCurve hostPipe = start.Owner as MEPCurve;
                Connector end = (from Connector c in hostPipe.ConnectorManager.Connectors
                                 where (int)c.ConnectorType == 1 && c.Id != start.Id
                                 select c).FirstOrDefault();
                XYZ dir = (start.Origin - end.Origin).Normalize();
                XYZ pipeHorizontalDirection = new XYZ(dir.X, dir.Y, 0.0).Normalize(); //Only for horizontal pipes! Fix this if the pipes are in any other direction
                XYZ connectorDirection = -capConnector.CoordinateSystem.BasisZ;
                double zRotationAngle = pipeHorizontalDirection.AngleTo(connectorDirection);
                Transform trf = Transform.CreateRotationAtPoint(XYZ.BasisZ, zRotationAngle, start.Origin);
                XYZ testRotation = trf.OfVector(connectorDirection).Normalize();
                if (Math.Abs(testRotation.DotProduct(pipeHorizontalDirection) - 1) > 0.00001)
                    zRotationAngle = -zRotationAngle;
                Line axis = Line.CreateBound(start.Origin, start.Origin + XYZ.BasisZ); //CREATE BOUND FOR ROTATION FFS!!!! It cost me two days of frustration
                cap.Location.Rotate(axis, zRotationAngle);
                #endregion

                #region Debug

                //var marker = Helper.PlaceAdaptiveMarkerLine("Green", start.Origin, start.Origin + XYZ.BasisZ);
                //Parameter parameter = cap.get_Parameter(BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS);
                //parameter.Set(Conversion.RadianToDegree(zRotationAngle).ToString());

                ////Determine the actual axis of rotation
                //XYZ newConnectorDirection = -capConnector.CoordinateSystem.BasisZ;
                //XYZ actualAxis = (connectorDirection.CrossProduct(newConnectorDirection)).Normalize();
                //var markerActualAxis = Helper.PlaceAdaptiveMarkerLine("Red", start.Origin, start.Origin + actualAxis);

                ////Marker line at original connector direction
                //Helper.PlaceAdaptiveMarkerLine("Yellow", start.Origin, start.Origin + pipeHorizontalDirection);

                ////Marker line at new connector direction
                //Helper.PlaceAdaptiveMarkerLine("Orange", start.Origin, start.Origin + connectorDirection);

                #endregion

                Parameter sizeParameter = cap.LookupParameter("Nominal Diameter 1"); //Hardcoded until inmplement
                sizeParameter.Set(pipe1.Diameter);

                elementSymbol.CreatedElement = cap;

                c1.ConnectTo(c2);

                return Result.Succeeded;

            }

            catch (Exception e)
            {
                Console.WriteLine(e);
                //throw new Exception(e.Message);
            }
            return Result.Failed;
        }

        public Result FLANGE(ElementSymbol elementSymbol)
        {
            try
            {
                //Place the instance at calculated midpoint
                Element flange = PCFImport.doc.Create.NewFamilyInstance(elementSymbol.CentrePoint.Xyz, elementSymbol.FamilySymbol, StructuralType.NonStructural);

                //Get all pipe connectors
                HashSet<Connector> allPipeConnectors = ch.GetAllPipeConnectors();

                //Get the actual endpoints of the flange
                XYZ p1 = elementSymbol.EndPoint1.Xyz; XYZ p2 = elementSymbol.EndPoint2.Xyz;

                //Get the flange connectors
                Connector c1 = ch.GetSecondaryConnector(ch.GetConnectorSet(flange));

                //Determine the corresponding pipe connectors
                var c2 = (from Connector c in allPipeConnectors where Util.IsEqual(p1, c.Origin) select c).FirstOrDefault();

                Pipe pipe1 = null;
                if (c2 != null) pipe1 = c2.Owner as Pipe;

                if (c2 == null)
                {
                    pipe1 = ch.CreateDummyPipe(p1, p2, elementSymbol.EndPoint1, elementSymbol);
                    c2 = ch.MatchConnector(p1, pipe1);
                    elementSymbol.DummyToDelete = pipe1;
                }

                ch.RotateElementInPosition(elementSymbol, c1, c2, flange);

                Parameter sizeParameter = flange.LookupParameter("Nominal Diameter 1"); //Hardcoded until inmplement
                sizeParameter.Set(pipe1.Diameter);

                elementSymbol.CreatedElement = flange;

                c1.ConnectTo(c2);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return Result.Succeeded;
        }

        public Result FLANGE_BLIND(ElementSymbol elementSymbol)
        {
            //Flange blind is suspended until after accessories are implemented
            //Because it attaches mostly to accessories in my model
            //Because of this it fails to find the placement point because accessories do not yet exist

            ////CAP can be reused
            //Result result = CAP(elementSymbol);
            //return result;
            return Result.Succeeded;
        }

        public Result REDUCER_CONCENTRIC(ElementSymbol elementSymbol)
        {
            try
            {
                //Get all pipe connectors
                HashSet<Connector> allPipeConnectors = ch.GetAllPipeConnectors();

                //Get the actual endpoints of the elbow
                XYZ p1 = elementSymbol.EndPoint1.Xyz; XYZ p2 = elementSymbol.EndPoint2.Xyz;
                //Determine the corresponding pipe connectors
                var c1 = (from Connector c in allPipeConnectors where Util.IsEqual(p1, c.Origin) select c).FirstOrDefault();
                var c2 = (from Connector c in allPipeConnectors where Util.IsEqual(p2, c.Origin) select c).FirstOrDefault();

                //Handle the missing connectors by creating dummy pipes

                Pipe pipe1 = null; Pipe pipe2 = null;

                if (c1 == null)
                {
                    pipe1 = ch.CreateDummyPipe(p1, elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint1, elementSymbol);
                    c1 = ch.MatchConnector(p1, pipe1);
                }

                if (c2 == null)
                {
                    pipe2 = ch.CreateDummyPipe(p2, elementSymbol.CentrePoint.Xyz, elementSymbol.EndPoint2, elementSymbol);
                    c2 = ch.MatchConnector(p2, pipe2);
                }

                if (c1 != null && c2 != null)
                {
                    Element element = PCFImport.doc.Create.NewTransitionFitting(c1, c2);
                    if (pipe1 != null) PCFImport.doc.Delete(pipe1.Id);
                    if (pipe2 != null) PCFImport.doc.Delete(pipe2.Id);
                    elementSymbol.CreatedElement = element;
                    PCFImport.doc.Regenerate();
                    return Result.Succeeded;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            //If this point is reached, something has failed
            return Result.Failed;
        }

        public Result OLET(ElementSymbol elementSymbol)
        {
            try
            {
                //Get all pipe connectors
                var allConnectors = ch.GetALLConnectors(pif._doc);

                //Determine the branch1point of the olet which is the connection to the olet pipe
                XYZ bp1 = elementSymbol.Branch1Point.Xyz;
                XYZ cp = elementSymbol.CentrePoint.Xyz;

                Pipe pipe1 = null;
                //Pipe pipe2 = null;

                pipe1 = ch.CreateDummyPipe(bp1, elementSymbol.CentrePoint.Xyz, elementSymbol.Branch1Point, elementSymbol);
                var c1 = ch.MatchConnector(bp1, pipe1);

                ////Determine the corresponding pipe connector to bp1
                //var c1 = (from Connector c in allConnectors where Util.IsEqual(bp1, c.Origin) select c).FirstOrDefault();
                ////Get the owner of the connector
                //var owner = c1.Owner;

                //Find the target pipe
                var filter = new ElementClassFilter(typeof(Pipe));
                var view3D = ch.Get3DView(pif._doc);
                var refIntersect = new ReferenceIntersector(filter, FindReferenceTarget.Element, view3D);
                ReferenceWithContext rwc = refIntersect.FindNearest(c1.Origin, c1.CoordinateSystem.BasisZ);
                var refId = rwc.GetReference().ElementId;
                var pipeToConnectInto = (MEPCurve)pif._doc.GetElement(refId);

                if (c1 != null)
                {
                    Element element = PCFImport.doc.Create.NewTakeoffFitting(c1, pipeToConnectInto);
                    if (pipe1 != null) PCFImport.doc.Delete(pipe1.Id);
                    //if (pipe2 != null) PCFImport.doc.Delete(pipe2.Id);
                    elementSymbol.CreatedElement = element;
                    PCFImport.doc.Regenerate();
                    return Result.Succeeded;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            //If this point is reached, something has failed
            return Result.Failed;
        }

        public Result VALVE(ElementSymbol elementSymbol)
        {

            try
            {
                #region Valve as AdaptiveElement

                //elementSymbol.CreatedElement = Helper.PlaceAdaptiveFamilyInstance("GenValve: Std", elementSymbol.EndPoint1.Xyz,
                //    elementSymbol.EndPoint2.Xyz);
                //Parameter sizeParameter = elementSymbol.CreatedElement.LookupParameter("Nominal Diameter");
                //sizeParameter.Set(elementSymbol.EndPoint1.Diameter);
                //return Result.Succeeded;

                #endregion

                //The strange symbol activation thingie...
                //See: http://thebuildingcoder.typepad.com/blog/2014/08/activate-your-family-symbol-before-using-it.html
                if (!elementSymbol.FamilySymbol.IsActive)
                {
                    elementSymbol.FamilySymbol.Activate();
                    PCFImport.doc.Regenerate();
                }

                //Place the instance at calculated midpoint
                Element valve = PCFImport.doc.Create.NewFamilyInstance(elementSymbol.CentrePoint.Xyz, elementSymbol.FamilySymbol, StructuralType.NonStructural);

                //Get all pipe connectors
                HashSet<Connector> allConnectors = ch.GetALLConnectors(valve.Document);

                //Get the actual endpoints of the valve
                XYZ p1 = elementSymbol.EndPoint1.Xyz; XYZ p2 = elementSymbol.EndPoint2.Xyz;

                //Get the primary connector of valve
                Connector c1 = ch.GetPrimaryConnector(ch.GetConnectorSet(valve));

                //Determine the corresponding connectors from all, if not found -- abort
                Connector c2 = (from Connector c in allConnectors where Util.IsEqual(p1, c.Origin) select c).FirstOrDefault() ??
                               (from Connector c in allConnectors where Util.IsEqual(p2, c.Origin) select c).FirstOrDefault();

                Pipe pipe1 = null;
                if (c2 != null) pipe1 = c2.Owner as Pipe;

                if (c2 == null)
                {
                    pipe1 = ch.CreateDummyPipe(p1, p2, elementSymbol.EndPoint1, elementSymbol);
                    c2 = ch.MatchConnector(p1, pipe1);
                    elementSymbol.DummyToDelete = pipe1;
                }

                ch.RotateElementInPosition(elementSymbol, c1, c2, valve);

                Parameter sizeParameter = valve.LookupParameter("Nominal Diameter"); //Hardcoded until inmplement
                sizeParameter.Set(pipe1.Diameter);

                Parameter lengthParameter = valve.LookupParameter("Length"); //Hardcoded until inmplement
                lengthParameter.Set(Util.Distance(p1, p2));

                elementSymbol.CreatedElement = valve;

                return Result.Succeeded;

            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                return Result.Failed;
                //throw new Exception("Valve creation generated following error: " + e.Message);
            }
        }
    }
}
