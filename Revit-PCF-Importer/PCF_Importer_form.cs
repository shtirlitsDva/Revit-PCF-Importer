using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BuildingCoder;
using Form = System.Windows.Forms.Form;
using mySettings = Revit_PCF_Importer.Properties.Settings;
using iv = PCF_Functions.InputVars;

namespace Revit_PCF_Importer
{
    public partial class PCF_Importer_form : Form
    {
        public static ExternalCommandData _commandData;
        public static UIApplication _uiapp;
        public static UIDocument _uidoc;
        public static Document _doc;
        public string _message;
        public string _excelPath = string.Empty;
        private string _pcfPath = string.Empty;
        public string[] readLines;

        public PCF_Importer_form(ExternalCommandData cData, ref string message)
        {
            InitializeComponent();
            _commandData = cData;
            _uiapp = _commandData.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;
            _message = message;

            //Init saved values
            _excelPath = mySettings.Default.excelPath;
            _pcfPath = mySettings.Default.pcfPath;

            //Init textboxes
            textBox1.Text = _pcfPath;
            textBox2.Text = _excelPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Get the file path
                _pcfPath = openFileDialog1.FileName;
                textBox1.Text = _pcfPath;
                //Set the PCF file path setting
                mySettings.Default.pcfPath = _pcfPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                //Get the file path
                _excelPath = openFileDialog2.FileName;
                textBox2.Text = _excelPath;
                //Set the PCF file path setting
                mySettings.Default.excelPath = _excelPath;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            PCFImport pcfImport = new PCFImport();
            Result result = pcfImport.ExecuteMyCommand(_uiapp, ref _message);
            if (result == Result.Succeeded) Util.InfoMsg("PCF data imported successfully!");
            if (result == Result.Failed) Util.InfoMsg("PCF data import failed for some reason.");
        }
    }
}
