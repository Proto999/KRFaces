using System;
using System.Windows.Forms;
using SolidEdgeCommunity;
using SolidEdgeFramework;
using SolidEdgeFrameworkSupport;
using SolidEdgePart;
using SolidEdgeConstants;
using SolidEdgeCommunity.Extensions;
using System.Threading;
using System.Runtime.InteropServices;
using System.Reflection;
using SolidEdgeGeometry;
using SolidEdgeAssembly;
using SolidEdgeDraft;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;

namespace Artem
{
    public partial class Form1 : Form
    {
        private SolidEdgeFramework.Application seApplication;
        private SolidEdgePart.PartDocument sePartDocument;
        private SolidEdgeGeometry.Faces faces;
        private SolidEdgeFramework.FaceStyles faceStyles;
        private SolidEdgeFramework.FaceStyle faceStyle;
        private int currentFaceIndex = 1;
        private Panel panel;
        private SolidEdgePart.RefPlanes refPlanes = null;
        private SolidEdgePart.RefPlane selectedPlane = null;
        private SolidEdgePart.Sketchs sketchs = null;
        private SolidEdgePart.Sketch sketch = null;
        private SolidEdgePart.Profiles profiles = null;
        private SolidEdgePart.Profile profile = null;
        private SolidEdgeFrameworkSupport.Lines2d lines2D = null;
        private SolidEdgePart.Sketch3D Sketch3D = null;
        private SolidEdgePart.Points3D Points3D = null;
        private object vertices = null;

        private int currentPlaneIndex = 1;
        private SolidEdgeFramework.SelectSet selectSet = null;

        public Form1()
        {
            InitializeComponent();

            try
            {
                OleMessageFilter.Register();

                seApplication = SolidEdgeUtils.Connect(true);

                if (seApplication.ActiveDocumentType != SolidEdgeFramework.DocumentTypeConstants.igPartDocument)
                {
                    MessageBox.Show("Please open a Part document.");
                    Close();
                    return;
                }

                panel = new Panel();
                panel.AutoScroll = true;
                panel.Location = new System.Drawing.Point(0, 0);
                panel.Size = new System.Drawing.Size(500, 300);
                Controls.Add(panel);

                Button btnNextFace = new Button();
                btnNextFace.Text = "Следующая грань";
                btnNextFace.Top = 20;
                btnNextFace.Left = 20;
                btnNextFace.Width = 200;
                btnNextFace.Click += btnNextFace_Click;
                panel.Controls.Add(btnNextFace);

                Button btnPreviousFace = new Button();
                btnPreviousFace.Text = "Предыдущая грань";
                btnPreviousFace.Top = 20;
                btnPreviousFace.Left = 250;
                btnPreviousFace.Width = 200;
                btnPreviousFace.Click += btnPreviousFace_Click;
                panel.Controls.Add(btnPreviousFace);

                Button btnNextPlane = new Button();
                btnNextPlane.Text = "Следующая плоскость";
                btnNextPlane.Top = 70;
                btnNextPlane.Left = 20;
                btnNextPlane.Width = 200;
                btnNextPlane.Click += btnNextPlane_Click;
                panel.Controls.Add(btnNextPlane);

                Button btnPreviousPlane = new Button();
                btnPreviousPlane.Text = "Предыдущая плоскость";
                btnPreviousPlane.Top = 70;
                btnPreviousPlane.Left = 250;
                btnPreviousPlane.Width = 200;
                btnPreviousPlane.Click += btnPreviousPlane_Click;
                panel.Controls.Add(btnPreviousPlane);

                Button btnGetSketch = new Button();
                btnGetSketch.Text = "Создать эскиз по грани";
                btnGetSketch.Top = 110;
                btnGetSketch.Left = 135;
                btnGetSketch.Width = 200;
                btnGetSketch.Click += btnGetSketch_Click;
                panel.Controls.Add(btnGetSketch);



                Button btnGetSketchPlane = new Button();
                btnGetSketchPlane.Text = "Создать эскиз по плоскости";
                btnGetSketchPlane.Top = 145;
                btnGetSketchPlane.Left = 135;
                btnGetSketchPlane.Width = 200;
                btnGetSketchPlane.Click += applySelectedPlane_Click;
                panel.Controls.Add(btnGetSketchPlane);

                sePartDocument = (PartDocument)seApplication.ActiveDocument;

                faces = (Faces)sePartDocument.Models.Item(1)
                    .ExtrudedProtrusions.Item(1)
                    .Faces[SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll];

                faceStyles = (FaceStyles)sePartDocument.FaceStyles;
                faceStyle = (FaceStyle)faceStyles.Item(3);

                ApplyFaceStyle(currentFaceIndex);

                selectSet = sePartDocument.SelectSet;
                refPlanes = GetRefPlanesFromActiveDocument();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Close();
            }
            finally
            {
                OleMessageFilter.Unregister();
            }
        }

        private void ApplyFaceStyle(int index)
        {
            for (int i = 1; i <= faces.Count; i++)
            {
                SolidEdgeGeometry.Face face = (SolidEdgeGeometry.Face)faces.Item(i);
                face.Style = (i == index) ? faceStyle : null;
            }
        }

        private void btnNextFace_Click(object sender, EventArgs e)
        {
            currentFaceIndex = Math.Min(currentFaceIndex + 1, faces.Count);
            ApplyFaceStyle(currentFaceIndex);
        }

        private void btnPreviousFace_Click(object sender, EventArgs e)
        {
            currentFaceIndex = Math.Max(currentFaceIndex - 1, 1);
            ApplyFaceStyle(currentFaceIndex);
        }

        private void btnNextPlane_Click(object sender, EventArgs e)
        {
            if (refPlanes != null)
            {
                selectSet.RemoveAll();
                currentPlaneIndex = Math.Min(currentPlaneIndex + 1, refPlanes.Count);
                selectedPlane = refPlanes.Item(currentPlaneIndex);
                //ApplySelectedPlane();
                selectSet.Add(selectedPlane);
            }
        }

        private void btnPreviousPlane_Click(object sender, EventArgs e)
        {
            if (refPlanes != null)
            {
                currentPlaneIndex = Math.Max(currentPlaneIndex - 1, 1);
                selectedPlane = refPlanes.Item(currentPlaneIndex);
                //ApplySelectedPlane();
                //UpdateInterfaceForCurrentPlane();
            }
        }

        private void applySelectedPlane_Click(object sender, EventArgs e)
        {
            ApplySelectedPlane();
        }

        private void ApplySelectedPlane()
        {
            if (selectedPlane != null)
            {
                CreateSketchOnSelectedPlane(selectedPlane);
                //UpdateInterfaceForCurrentPlane();
            }
        }

        private void CreateSketchOnSelectedPlane(SolidEdgePart.RefPlane plane)
        {
            try
            {
                OleMessageFilter.Register();

                seApplication.DoIdle();

                sketchs = sePartDocument.Sketches;
                sketch = sketchs.Add();
                profiles = sketch.Profiles;
                profile = profiles.Add(plane);
                lines2D = profile.Lines2d;

                lines2D.AddBy2Points(0, 0, 1, 0);
                lines2D.AddBy2Points(1, 0, 1, 1);
                lines2D.AddBy2Points(1, 1, 0, 1);
                lines2D.AddBy2Points(0, 1, 0, 0);

                profile.End(SolidEdgePart.ProfileValidationType.igProfileClosed);

                System.Threading.Thread.Sleep(500);


                seApplication.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewISOView);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }
        }

        private SolidEdgePart.RefPlanes GetRefPlanesFromActiveDocument()
        {
            SolidEdgePart.RefPlanes refPlanes = null;

            try
            {
                refPlanes = sePartDocument.RefPlanes;

                if (refPlanes == null || refPlanes.Count == 0)
                {
                    MessageBox.Show("В документе нет плоскостей.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при получении плоскостей: {ex.Message}");
            }

            return refPlanes;
        }

        public static void CreatePlaneRelativeToFaceUsingCoordinateSystem(SolidEdgePart.PartDocument partDocument, int faceIndex, RefPlanes refPlanes, int currentFaceIndex, Faces faces, Face face)
        {
            Models models = partDocument.Models;
            Model model = models.Item(1);
            // Get the current selected face
            SolidEdgeGeometry.Face selectedFace = (SolidEdgeGeometry.Face)faces.Item(currentFaceIndex);
            CoordinateSystems coordinateSystems1 = partDocument.CoordinateSystems;
            //CoordinateSystems coordinateSystems = model.CoordinateSystems;
            CoordinateSystem coordinateSystem = coordinateSystems1.Item(1); // Выбираем первую систему координат

            double centerX, centerY, centerZ;
            //face.GetParamOnFace(out centerX, out centerY, out centerZ);

            double normalX, normalY, normalZ;
           // face.GetNormal(out normalX, out normalY, out normalZ);

            //RefPlane refPlane = refPlanes.AddRelativeToCoordinateSystem(centerX, centerY, centerZ, normalX, normalY, normalZ, coordinateSystem);
        }

        private void btnGetSketch_Click(object sender, EventArgs e)
        {
            try
            {
                OleMessageFilter.Register();

                if (selectedPlane != null)
                {
                    // Clear the previously selected plane
                    selectedPlane.Delete();
                    selectedPlane = null;
                }

                // Get the current selected face
                SolidEdgeGeometry.Face selectedFace = (SolidEdgeGeometry.Face)faces.Item(currentFaceIndex);

                // Получение ребер грани
                SolidEdgeGeometry.Edges edges = selectedFace.Edges as SolidEdgeGeometry.Edges;

                // Проверка наличия ребер и создание плоскости
                if (edges != null && edges.Count > 0)
                {
                    // Выбираем первое ребро для определения направления плоскости
                    SolidEdgeGeometry.Edge firstEdge = edges.Item(1) as SolidEdgeGeometry.Edge;

                    // Проверка наличия первого ребра
                    if (firstEdge != null)
                    {
                        // Получаем координаты первой вершины первого ребра
                        double[] vertexCoordinates = firstEdge.StartVertex as double[];

                        // Определение точки на новой плоскости (выбранной вершиной)
                        double[] origin = vertexCoordinates;

                        // Расстояние для плоскости, которое можно настроить в соответствии с вашими потребностями
                        double distance = 0.01; // Замените на необходимое расстояние

                        // Получаем направление плоскости из нормали первой вершины
                        double[] normal = firstEdge.StartVertex as double[];

                        // Создаем точку для опоры
                        double[] pivotPoint = new double[] { 1.0, 2.0, 3.0 };
                        double[] pivotOrigin = new double[] { 0.0, 0.0, 0.0 }; // Например, центр масс объекта
                        bool flipNormal = false; // или true, в зависимости от требований

                        if (normal != null && normal.Length >= 3)
                        {
                            // Создание новой плоскости параллельной грани с использованием расстояния
                            SolidEdgePart.RefPlane parallelPlane = sePartDocument.RefPlanes.AddParallelByTangent(
                                ParentPlane: refPlanes.Item(1),
                                selectedFace,
                                TangentPositionFlag: SolidEdgePart.KeyPointExtentConstants.igTangentNormal,
                                Pivot: pivotPoint,
                                PivotOrigin: pivotOrigin,
                                Local: null
                                                         );

                            // Теперь у вас есть новая плоскость, параллельная выбранной грани
                        }
                            }
                        }
                    
               
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                OleMessageFilter.Unregister();
            }
            }
        }
}

