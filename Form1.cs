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

                

                // Получим текущую выбранную грань
                SolidEdgeGeometry.Face selectedFace = (SolidEdgeGeometry.Face)faces.Item(currentFaceIndex);

                // Создадим плоскость относительно текущей выбранной грани
                selectedPlane = sePartDocument.RefPlanes.AddParallelByDistance(
                    selectedFace,
                    0.0001,
                    SolidEdgePart.ReferenceElementConstants.igNormalSide,
                    false,
                    Missing.Value,
                    Missing.Value,
                    Missing.Value);

                
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
