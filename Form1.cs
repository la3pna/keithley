using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NationalInstruments.VisaNS;
using System.IO;
using System.Globalization;


namespace keithley
{
    public partial class Form1 : Form
    {

        private MessageBasedSession mbSession;
        public static string strVISArsrc;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                string[] stSesssion_avaible = ResourceManager.GetLocalManager().FindResources("?*");
                comboBox1.Items.AddRange(stSesssion_avaible);
            }
            catch (Exception exp) { MessageBox.Show(exp.Message + "\n \n This error may be due to VISA not being installed or the instrument not found"); }
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            strVISArsrc = comboBox1.Text;
            try
            {
                mbSession = (MessageBasedSession)ResourceManager.GetLocalManager().Open(strVISArsrc);
                // if (mbSession.ResourceManufacturerID == ){
                //  string a  = (string)mbSession.ResourceManufacturerID ;
                // }
                // here it should check for the appropriate Rigol or Agilent scope, given that agilent have same data set


            }
            catch (InvalidCastException)
            {
                MessageBox.Show("Resource selected is not an message based session");
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            List<float> lista = new List<float>();
            List<float> listb = new List<float>();
            List<float> listc = new List<float>();
            List<float> listd = new List<float>();
            try
            {
              mbSession.Write("smua.OUTPUT_DCVOLTS");
              //  mbSession.Write("smua.source.rangev = 5");
               
                for(int i=1;i<102;i++){
                float levelv = (0.001f * i) - 0.051f;
                string mbstring = "smua.source.levelv=" + levelv.ToString("0.###") ;
                mbSession.Write(mbstring);
                mbSession.Write("smua.OUTPUT_ON");
                string voltage = mbSession.Query("smua.measure.i()");
                mbSession.Write("smua.OUTPUT_OFF");
                lista.Add(levelv);
                listb.Add(Convert.ToSingle(voltage));
             
                }



                  for (int i = 0; i<101; i++)
            {
                float levelv = ( i) * 0.02f;
                string mbstring = "smua.source.levelv=" + levelv.ToString("0.###") ;
                mbSession.Write(mbstring);
                mbSession.Write("smua.OUTPUT_ON");
                string voltage = mbSession.Query("smua.measure.i()");
                mbSession.Write("smua.OUTPUT_OFF");
                listc.Add(levelv);
                listd.Add(Convert.ToSingle(voltage));
      
                }


                  float[] aArrayn = lista.ToArray();
                  float[] bArrayn = listb.ToArray();
                  float[] cArrayn = listc.ToArray();
                  float[] dArrayn = listd.ToArray();


                  Stream myStream;
                  SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                  saveFileDialog1.Filter = "Comma separated (*.csv)|*.csv";
                  saveFileDialog1.FilterIndex = 1;
                  saveFileDialog1.RestoreDirectory = true;

                  if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                  {
                      if ((myStream = saveFileDialog1.OpenFile()) != null)
                      {
                          var extension = Path.GetExtension(saveFileDialog1.FileName);
                          switch (extension.ToLower())
                          {
                              case ".csv":
                                  StreamWriter wText = new StreamWriter(myStream);
                                  wText.WriteLine(textBox1.Text);
                                 
                                  for (int i = 0; i <= aArrayn.Length - 1; i++)
                                  {
                                      wText.WriteLine(Convert.ToString(aArrayn[i], CultureInfo.InvariantCulture) + ',' + Convert.ToString(bArrayn[i], CultureInfo.InvariantCulture) + "," + Convert.ToString(cArrayn[i], CultureInfo.InvariantCulture) + ',' + Convert.ToString(dArrayn[i], CultureInfo.InvariantCulture));
                                  }
                                  wText.Flush();
                                  wText.Close();
                                  break;
                              default:
                                  throw new ArgumentOutOfRangeException(extension);
                          }
                      }
                  }
      


                 }
                 catch(Exception ex)
                 {}



            }

        private void label2_Click(object sender, EventArgs e)
        {

        }





        





    }
}
