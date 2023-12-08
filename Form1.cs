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
        object vertices = null;

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
                panel.AutoScroll = false;
                panel.Location = new System.Drawing.Point(0, 0);
                panel.Size = new System.Drawing.Size(500, 100);
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

                Button btnGetSketch = new Button();
                btnGetSketch.Text = "Создать эскиз";
                btnGetSketch.Top = 50;
                btnGetSketch.Left = 135;
                btnGetSketch.Width = 200;
                btnGetSketch.Click += btnGetSketch_Click;
                panel.Controls.Add(btnGetSketch);

                sePartDocument = (PartDocument)seApplication.ActiveDocument;

                faces = (Faces)sePartDocument.Models.Item(1)
                    .ExtrudedProtrusions.Item(1)
                    .Faces[SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll];

                faceStyles = (FaceStyles)sePartDocument.FaceStyles;
                faceStyle = (FaceStyle)faceStyles.Item(3);

                ApplyFaceStyle(currentFaceIndex);
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

                // Get the vertices of the selected face
                object[] verticesArray = (object[])selectedFace.Vertices;

                // Check if there are at least three vertices
                if (verticesArray.Length < 3)
                {
                    Console.WriteLine("Ошибка: Недостаточно вершин для создания плоскости.");
                    return;
                }

                // Получение координат первых трех вершин
                double[] point1 = ((double[])verticesArray[0]).ToArray();
                double[] point2 = ((double[])verticesArray[1]).ToArray();
                double[] point3 = ((double[])verticesArray[2]).ToArray();

                // Отладочные выводы
                //Debug.WriteLine($"point1: X={point1[0]}, Y={point1[1]}, Z={point1[2]}");
                //Debug.WriteLine($"point2: X={point2[0]}, Y={point2[1]}, Z={point2[2]}");
                //Debug.WriteLine($"point3: X={point3[0]}, Y={point3[1]}, Z={point3[2]}");


                // Создание координатной системы, проходящей через три точки
                //CoordinateSystems coordinateSystems = (CoordinateSystems)sePartDocument.CoordinateSystems;
                //CoordinateSystem coordinateSystem = coordinateSystems.AddBy3Points(
                // point1[0], point1[1], point1[2],
                // point2[0], point2[1], point2[2],
                // point3[0], point3[1], point3[2]
                // );

                // Получение плоскости относительно координатной системы
                //selectedPlane = (RefPlane)sePartDocument.RefPlanes.AddByCoordinateSystem(
                //coordinateSystem
                //);

                if (selectedPlane == null)
                {
                    Console.WriteLine("Ошибка: Плоскость не создана.");
                }
                else
                {
                    // Отображение результата в Solid Edge с использованием методов управления видимостью
                    seApplication.StartCommand(SolidEdgeConstants.PartCommandConstants.PartViewFit);
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