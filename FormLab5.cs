using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Windows.Forms;

namespace lab7
{
    public partial class FormLab5 : Form
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private SelectionMgr swSelMgr;
        private double L1 = 0, L2 = 0, L3 = 0, L4 = 0, L5 = 0, L6 = 0, L7 = 0, L8 = 0, L9 = 0, L10 = 0;
        private string front = "Спереди", top = "Сверху", left = "Слева", right = "Справа", bottom = "Снизу";
        private SketchManager skm;
        private Feature footing;
        private Feature feature;
        private Feature addFeature;
        private Entity ent;
        private const double changeUnit = 1000;
        private int i = 0;
        private bool res;

        public FormLab5()
        {
            InitializeComponent();
        }

        private void FormLab5_Load(object sender, EventArgs e)
        {
            swApp = new SldWorks();
            swApp.Visible = true;
        }

        private void DrawFullModel(object sender, EventArgs e)
        {
            checkDrawing();
            ReadVaribles();
            for (int j = 0; j < 5; j++)
            {
                DrawOneStep(null, null);
                if (!res)
                {
                    break;
                }
            }
            Close();
        }

        private bool ReadVaribles()
        {
            Double.TryParse(textBox1.Text, out L1);
            L1 /= changeUnit;

            Double.TryParse(textBox2.Text, out L2);
            L2 /= changeUnit;

            Double.TryParse(textBox3.Text, out L3);
            L3 /= changeUnit;

            Double.TryParse(textBox4.Text, out L4);
            L4 /= changeUnit;

            Double.TryParse(textBox5.Text, out L5);
            L5 /= changeUnit;

            Double.TryParse(textBox6.Text, out L6);
            L6 /= changeUnit;

            Double.TryParse(textBox7.Text, out L7);
            L7 /= changeUnit;

            Double.TryParse(textBox8.Text, out L8);
            L8 /= changeUnit;

            Double.TryParse(textBox9.Text, out L9);
            L9 /= changeUnit;

            Double.TryParse(textBox10.Text, out L10);
            L10 /= changeUnit;

            try
            {
                if (L1 == 0 || L2 == 0 || L3 == 0 || L4 == 0 || L5 == 0 || L6 == 0 || L7 == 0 || L9 == 0)
                {
                    throw new Exception("Длина не может быть равна 0");
                }

                if (L9 > L1)
                {
                    throw new Exception("L1 должно быть больше чем L9");
                }

                if (L3 >= L2)
                {
                    throw new Exception("L2 должно быть больше чем L3");
                }

                if (L4 <= L7)
                {
                    throw new Exception("L4 Должно быть больше чем L7");
                }

                if (L5 >= L1)
                {
                    throw new Exception("L1 Должно больше чем L5");
                }

                if (L10 >= L1)
                {
                    throw new Exception("L1 Должно больше чем L10");
                }

                if (L6 + L10 > L1)
                {
                    throw new Exception("L6+L10 не может быть больше чем L1");
                }

                if (L5 <= L6 / 2)
                {
                    throw new Exception("Невозможно разместить вырез");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }

            return true;
        }

        private void Switch()
        {
            textBox1.Enabled = !textBox1.Enabled;
            textBox2.Enabled = !textBox2.Enabled;
            textBox3.Enabled = !textBox3.Enabled;
            textBox4.Enabled = !textBox4.Enabled;
            textBox5.Enabled = !textBox5.Enabled;
            textBox6.Enabled = !textBox6.Enabled;
            textBox7.Enabled = !textBox7.Enabled;
            textBox8.Enabled = !textBox8.Enabled;
            textBox9.Enabled = !textBox9.Enabled;
            textBox10.Enabled = !textBox10.Enabled;
        }

        private void SelectPlane(string name)
        {
            swModel.Extension.SelectByID2(name, "PLANE", 0, 0, 0, false, 0, null, 0);
        }

        private void checkDrawing()
        {
            if (swApp.ActiveDoc == null)
            {
                swModel = (ModelDoc2)swApp.INewPart();
                swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
                skm = swModel.SketchManager;
            }
            else
            {
                swModel = (ModelDoc2)swApp.ActiveDoc;
                swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
                skm = swModel.SketchManager;
            }
        }

        private void checkBoxDfltParams_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxDfltParams.Checked)
            {
                textBox1.ReadOnly = true;
                textBox1.Text = "95";
                textBox2.ReadOnly = true;
                textBox2.Text = "85";
                textBox3.ReadOnly = true;
                textBox3.Text = "50";
                textBox4.ReadOnly = true;
                textBox4.Text = "65";
                textBox5.ReadOnly = true;
                textBox5.Text = "65";
                textBox6.ReadOnly = true;
                textBox6.Text = "25";
                textBox7.ReadOnly = true;
                textBox7.Text = "40";
                textBox8.ReadOnly = true;
                textBox8.Text = "20";
                textBox9.ReadOnly = true;
                textBox9.Text = "55";
                textBox10.ReadOnly = true;
                textBox10.Text = "50";
            }
            else
            {
                textBox1.ReadOnly = false;
                textBox2.ReadOnly = false;
                textBox3.ReadOnly = false;
                textBox4.ReadOnly = false;
                textBox5.ReadOnly = false;
                textBox6.ReadOnly = false;
                textBox7.ReadOnly = false;
                textBox8.ReadOnly = false;
                textBox9.ReadOnly = false;
                textBox10.ReadOnly = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            swModel = (ModelDoc2)swApp.ActiveDoc;
            swModel.SetUnits((short)swLengthUnit_e.swMM, (short)swFractionDisplay_e.swDECIMAL, 0, 0, false);
            skm = swModel.SketchManager;

            object[] swFeatures = swModel.FeatureManager.GetFeatures(true);

            foreach (var ftr in swFeatures)
            {
                Feature swFtr = ftr as Feature;
                if (swFtr.Name.Contains("Бобышка") || swFtr.Name.Contains("Вырез"))
                {
                    swFtr.Select2(true, -1);
                }
            }
            swModel.EditDelete();

            foreach (var ftr in swFeatures)
            {
                Feature swFtr = ftr as Feature;
                if (swFtr.Name.Contains("Эскиз"))
                {
                    swFtr.Select2(true, -1);
                }
            }
            swModel.EditDelete();

            button1.Enabled = true;
            button3.Enabled = false;
        }

        private void DrawOneStep(object sender, EventArgs e)
        {
            checkDrawing();
            swApp.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swInputDimValOnCreate, false);
            swModel.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDisplayAnnotations, true);
            if (i == 0)
            {
                res = ReadVaribles();
                if (!res)
                {
                    return;
                }
                Switch();
            }

            switch (i)
            {
                case 0:
                    Base();
                    i++;
                    break;

                case 1:
                    BackSide();
                    i++;
                    break;

                case 2:
                    BackSideСutout();
                    i++;
                    break;

                case 3:
                    Ribs();
                    i++;
                    break;

                case 4:
                    FrontСutout();
                    button1.Enabled = false;
                    button3.Enabled = true;
                    Switch();
                    i = 0;
                    break;
            }
        }

        #region Building

        private void Base()
        {
            SelectPlane(right);
            var rect = skm.CreateCenterRectangle(0, 0, 0, L1 / 2, (L4 - L7) / 2, 0);
            ((SketchSegment)rect[0]).Select(false);
            swModel.AddDimension(0, -(L4 - L7) / 2 - 15.0 / changeUnit, 0);

            footing = featureExtrusionBoth(L2);
            swModel.ClearSelection();
        }

        private void BackSide()
        {
            var faces = footing.GetFaces();
            ent = faces[4] as Entity;
            ent.Select(false);
            skm.InsertSketch(false);
            double x0 = 0.0 - L2 / 2;
            double y0 = 0.0 - L1 / 2;
            var rect = skm.CreateCornerRectangle(x0, y0, 0, x0 + L2, y0 + L6, 0);
            ((SketchSegment)rect[3]).Select(false);
            swModel.AddDimension(0.0 - L2 / 2 - 10.0 / changeUnit, 0, 0.0 - L1 / 2 + L6 / 2);

            feature = featureExtrusionBack(L4);
            swModel.ClearSelection();
        }

        private void BackSideСutout()
        {
            var faces = feature.GetFaces();
            ent = faces[0] as Entity;
            ent.Select(false);
            skm.InsertSketch(false);

            if (L8 != 0)
            {
                double x0 = 0.0 - L5 / 2;
                double y0 = 0.0 + L1 / 2;
                var rect = skm.CreateCornerRectangle(x0, y0, 0, x0 + L5, y0 - L6, 0);
                ((SketchSegment)rect[0]).Select(false);
                swModel.AddDimension(0.0, 0, 0 - L1 / 2 - 10.0 / changeUnit);

                feature = featureCutDepthFront(L8);
            }
            swModel.ClearSelection();
        }

        private void Ribs()
        {
            swModel.ClearSelection();
            var faces = footing.GetFaces();
            ent = faces[1] as Entity;
            ent.Select(false);
            skm.InsertSketch(false);

            double x0 = 0.0 - L1 / 2 + L6;
            double y0 = 0.0 + L4 / 2 + L7 / 2;
            if (L8 != 0)
            {
                var line1 = skm.CreateLine(-L1 / 2, (L4 - L7) / 2, 0, -L1 / 2 + L9, (L4 - L7) / 2, 0);
                line1.Select(false);
                swModel.AddDimension(1, -(L4 - L7) / 2 - 10.0 / changeUnit, -L1 / 2 + L9 / 2);
            }
            else
            {
                ent.Select(false);
                skm.InsertSketch(false);
            }

            var line = skm.CreateLine(x0, y0, 0, x0 - L6 + L9, y0 - L7, 0);
            line.Select(false);
            swModel.FeatureManager.InsertRib(false, true, (L2 - L5) / 2, 0, false, true, false, 0, false, true);

            ent = faces[0] as Entity;
            ent.Select(false);
            skm.InsertSketch(false);

            x0 = 0.0 + L1 / 2 - L6;
            y0 = 0.0 + L4 / 2 + L7 / 2;
            line = skm.CreateLine(x0, y0, 0, x0 + L6 - L9, y0 - L7, 0);
            line.Select(false);
            swModel.FeatureManager.InsertRib(false, true, (L2 - L5) / 2, 0, true, true, false, 0, false, true);

            swModel.ClearSelection();
        }

        private void FrontСutout()
        {
            if (L10 != 0)
            {
                var faces = footing.GetFaces();
                ent = faces[3] as Entity;
                ent.Select(false);
                skm.InsertSketch(false);
                var rect = skm.CreateCenterRectangle(0, 0, 0, L3 / 2, (L4 - L7) / 2, 0);
                ((SketchSegment)rect[0]).Select(false);
                swModel.AddDimension(0, (L4 - L7) / 2 + 10.0 / changeUnit, 0);

                featureCutDepthFront(L10);
            }
            swModel.ClearSelection();
        }

        #endregion Building

        #region Features

        private Feature featureExtrusionBoth(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(false, true, false,
            (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
            depth / 2, depth / 2, false, false, false, false, 0, 0, false, false, false, false, true,
            true, true, 0, 0, false);
        }

        private Feature featureExtrusionBack(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(true, false, true,
            (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
            depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
            true, true, 0, 0, false);
        }

        private Feature featureExtrusionFront(double depth)
        {
            return swModel.FeatureManager.FeatureExtrusion2(true, false, false,
            (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
            depth, 0, false, false, false, false, 0, 0, false, false, false, false, true,
            true, true, 0, 0, false);
        }

        private Feature featureCutDepthFront(double depth)
        {
            return swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature featureCutDepthBack(double depth)
        {
            return swModel.FeatureManager.FeatureCut2(true, true, false, (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                depth, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        private Feature featureCutThrough()
        {
            return swModel.FeatureManager.FeatureCut2(true, false, false, (int)swEndConditions_e.swEndCondThroughAllBoth, (int)swEndConditions_e.swEndCondThroughAllBoth,
                0, 0, false, false, false, false, 0, 0, false, false, false, false, false,
                false, false, false, false, false);
        }

        #endregion Features
    }
}