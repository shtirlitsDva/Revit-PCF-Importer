using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BuildingCoder;
using PCF_Functions;
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
        //Declare static dictionary for parsing
        public static PCF_Dictionary PcfDict;
        //Declare static dictionary for creating
        public static PCF_Creator PcfCreator;

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
            textBox3.Text = mySettings.Default.numberOfPipelinesDetected.ToString();

            //Init radiobuttons
            iv.ConfigureAll = mySettings.Default.configureAll;
            if (iv.ConfigureAll)
            {
                radioButton1.Checked = true;
                radioButton2.Checked = false;
            }
            else
            {
                radioButton1.Checked = false;
                radioButton2.Checked = true;
            }

            PcfDict = new PCF_Dictionary(new KeywordProcessor());
            PcfCreator = new PCF_Creator(new ProcessElements());
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
                //Read the file, get amount of pipelines and display the number
                FileReader fileReader = new FileReader();
                string[] readFile = fileReader.ReadFile(mySettings.Default.pcfPath);
                //This method collects all top-level element strings and creates ElementSymbols with data
                ElementCollection elements = new ElementCollection();
                Parser.CreateInitialElementList(elements, readFile);

                int numberOfPipelines =
                    (from ElementSymbol es in elements.Elements where es.ElementType == "PIPELINE-REFERENCE" select es)
                        .Count();

                textBox3.Text = numberOfPipelines.ToString();
                mySettings.Default.numberOfPipelinesDetected = numberOfPipelines;
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

        private void button4_Click(object sender, EventArgs e)
        {
            ElementCollection ExtractedElementCollection = new ElementCollection();
            PCF_Dictionary PcfDict = new PCF_Dictionary(new KeywordProcessor());

            //Read the input PCF file
            FileReader fileReader = new FileReader();
            string[] readFile = fileReader.ReadFile(mySettings.Default.pcfPath);
            //This method collects all top-level element strings and creates ElementSymbols with data
            Parser.CreateInitialElementList(ExtractedElementCollection, readFile);
            //This method compares all element symbols and gets the amount of line for their definition
            Parser.IndexElementDefinitions(ExtractedElementCollection, readFile);
            //This method extracts element data from the file
            Parser.ExtractElementDefinition(ExtractedElementCollection, readFile);
            //This method processes elements
            foreach (ElementSymbol elementSymbol in ExtractedElementCollection.Elements)
            {
                PcfDict.ProcessTopLevelKeywords(elementSymbol);
            }

            IList<ElementSymbol> elementList = (from ElementSymbol es in ExtractedElementCollection.Elements
                                               where !(
                                               string.Equals(es.PipelineReference, "PRE-PIPELINE") ||
                                               string.Equals(es.PipelineReference, "MATERIALS")
                                               )
                                               select es).ToList();

            if (iv.ConfigureAll == true)
            {
                PCF_Configuration.ExportAllConfigurationToExcel(elementList);
            }
            else
            {
                PCF_Configuration.ExportByPipelineConfigurationToExcel(elementList);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                iv.ConfigureAll = true;
                mySettings.Default.configureAll = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                iv.ConfigureAll = false;
                mySettings.Default.configureAll = false;
            }
        }
    }
}
