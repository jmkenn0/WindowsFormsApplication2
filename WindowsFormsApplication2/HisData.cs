using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using TxtFileConvert;

namespace WindowsFormsApplication2
{
    public partial class HisData : Form
    {

        public static DataTable fileArray = new DataTable();

        public HisData()
        {
            InitializeComponent();

            DataColumn InputColumn = new DataColumn();
            InputColumn.DataType = System.Type.GetType("System.String");
            InputColumn.ColumnName = "FileName";
            
            fileArray.Columns.Add(InputColumn);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Stream myStream = null;
            openFileDialog1.InitialDirectory = "c:\\";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }


            }

            //DataTable outputtable = Form1.OutputDataTable1();
            //loop through selection populating global string array varaible
            
            DataRow fileArrayRow;

            foreach (String file in openFileDialog1.FileNames)
            {

                fileArrayRow = fileArray.NewRow();
                fileArrayRow[0] = file;
                fileArray.Rows.Add(fileArrayRow);
                //MessageBox.Show(fileArray.Rows[0][0].ToString());
            }

            //stop here

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(fileArray.Rows[0][0].ToString());


            //Bring in TXT Sales File
            StreamReader streamreaderTXT;
            DataRow ExtractRow;
            int outputItem = 0;
            int outputItemPrev = 0;
            int counter = 0;
            

            //identify "tab" as the column delimiter in the txt file
             char[] delimiter = new char[] { ',' , ';', '\t'};
           // char[] delimiter = new char[] { '\t', ',' };
            //initialize output table


            DataTable fileHeaderRows = new DataTable();
            DataColumn fileHeaderColumn = new DataColumn();
            fileHeaderColumn.DataType= System.Type.GetType("System.String");
            fileHeaderColumn.ColumnName = "File";
            fileHeaderRows.Columns.Add(fileHeaderColumn);

            fileHeaderColumn = new DataColumn();
            fileHeaderColumn.DataType = System.Type.GetType("System.String");
            fileHeaderColumn.ColumnName = "Header Columns";
            fileHeaderRows.Columns.Add(fileHeaderColumn);

            DataRow fileHeaderRow = fileHeaderRows.NewRow();
            string path="";

            using (DataTable dtx = OutputDataTable1())
            {

                foreach (DataRow dt in fileArray.Rows)
                {
                    label1.Text = Path.GetFileNameWithoutExtension(dt[0].ToString());

                    path = Path.GetFileNameWithoutExtension(dt[0].ToString());
                    streamreaderTXT = new StreamReader(dt[0].ToString());
                    
                    
                    outputItem = 0;
                    outputItemPrev = 0;
                    string version = "";
                    string versionPrev = "";
                    string prevCMScheme = "";
                    string prevKey = "";
                    ExtractRow = dtx.NewRow();

                    //bool firstline = true;

                    while (!streamreaderTXT.EndOfStream)
                    {
                        
                        string[] inputstring = streamreaderTXT.ReadLine().Split(delimiter);

                        //string inputstring = streamreaderTXT.ReadLine();
                        //skip if beginning of file with the header info if it exists

                        //MessageBox.Show(inputstring[6].ToString());
                        if (inputstring[6].Substring(0,3)  != "CM1")
                            continue;
                        /*if (firstline)
                        {
                            fileHeaderRow[0] = dt[0].ToString();
                            fileHeaderRow[1] = inputstring.ToString();
                            fileHeaderRows.Rows.Add(fileHeaderRow);

                            fileHeaderRow = fileHeaderRows.NewRow();
                            //MessageBox.Show(dt[0].ToString() + "-" + inputstring.ToString());
                            firstline = false;
                             continue;
                        } */

                        
                        //initiate each field to "0" - Qlik requirement
                       /* for (int i = 0; i < ExtractRow.Table.Columns.Count; i++)
                        {

                            ExtractRow[i] = "0";

                        } */

                        //convert item into integer for comparion on whether or not to create a new row
                       // outputItem = convertItem(inputstring[6]);
                       // version = inputstring[3].ToString();

                        //if outputItemPrev=0 (indicating new file) or outputItem is less than or equal to outputItemPrev (indicating new row)
                        //create new row
                        if (prevKey=="" || inputstring[0] != prevKey)
                        {

                            if (prevKey!="")
                            {
                                //dtx.Rows.Add(ExtractRow);
                                ExtractRow = dtx.NewRow();
                            }
                            //cut and paste
                            //cut and paste
                            //take # in column 41, and post it in the corresponding contribution margin column in the new row
                            


                            
                           
                            ExtractRow["BURKS - Company Code"] = inputstring[9].ToString();  //need to figure this piece out?
                            ExtractRow["VERSI - Version"] = inputstring[3].ToString();
                            ExtractRow["CPLYEAR - Planning Year"] = inputstring[4].ToString();
                            
                            ExtractRow["CCOMPANY - Company"] = inputstring[8].ToString();
                            //not available
                            ExtractRow["VBUND - Partner Company"] = inputstring[18].ToString();
                            ExtractRow["CFCH002620 - Item"] = "";
                            ExtractRow["KOKRS - Controlling Area"] = inputstring[10].ToString();
                            ExtractRow["PRCTR - Profit Center"] = inputstring[13].ToString();
                            ExtractRow["CPPRCTR - Partner Profit Center"] = inputstring[11].ToString();
                            ExtractRow["WWBRN - Branch/Industry"] = inputstring[14].ToString();
                            ExtractRow["WWBUN - BU"] = inputstring[15].ToString();
                            ExtractRow["WWPRG - Product Group"] = inputstring[36].ToString();
                            ExtractRow["WWPPG - Partner Product Group"] = inputstring[37].ToString();
                            ExtractRow["WWART - Material Number"] = inputstring[46].ToString();//
                            ExtractRow["KUNWE - Ship-To (local)"] = inputstring[19].ToString();
                            ExtractRow["KNDNR - Sold-To (local)"] = inputstring[28].ToString();
                            ExtractRow["KUNRE - Bill-To (local)"] = inputstring[32].ToString();
                            ExtractRow["KUNRG - Payer (local)"] = inputstring[34].ToString();
                            ExtractRow["WWKUN - Ship-To Final (local)"] = inputstring[24].ToString();
                            ExtractRow["WWLWE - Country (Ship-To)"] = inputstring[20].ToString();
                            ExtractRow["LAND1 - Country (Sold-To)"] = inputstring[29].ToString();
                            ExtractRow["WWLRE - Country (Bill-To)"] = inputstring[33].ToString();
                            ExtractRow["WWLRG - Country (Payer)"] = inputstring[35].ToString();
                            ExtractRow["WWFCU - Country (Ship-To Final)"] = inputstring[25].ToString();

                            ExtractRow["KSTRG - Cost object"] = inputstring[60].ToString();
                            ExtractRow["WWKAF - Sales order"] = inputstring[64].ToString();
                            ExtractRow["KDPOS - Sales order item"] = inputstring[58].ToString();
                            ExtractRow["WWREN - CF invoice number"] = inputstring[63].ToString();
                            ExtractRow["CFCH00056 - Bill. Item"] = inputstring[59].ToString();//here
                           
                            ExtractRow["WWPST - Product Structure"] = inputstring[45].ToString();
                            ExtractRow["WWIDS - Identstring"] = inputstring[44].ToString();
                            ExtractRow["WWPRS - Product segment"] = inputstring[38].ToString();
                            ExtractRow["WWHWK - Product characteristic"] = inputstring[39].ToString();
                            ExtractRow["WWFTT - Lacquering"] = inputstring[40].ToString();
                            ExtractRow["WWKAS- Coating"] = inputstring[41].ToString();
                            ExtractRow["WWDRU - Print"] = inputstring[42].ToString();
                            ExtractRow["WWEND - Final Form"] = inputstring[43].ToString();
                            ExtractRow["WWBRA - Brand"] = inputstring[55].ToString();
                            ExtractRow["MATKL - Material Group"] = inputstring[50].ToString();
                            ExtractRow["BUDAT - Posting Date"] = convertDate(inputstring[71].ToString(), inputstring[72].ToString());
                            ExtractRow["FDAT - Invoice Date"] = convertDate(inputstring[71].ToString(), inputstring[72].ToString());
                            ExtractRow["FRWAE - Local Currency"] = //inputstring[76].ToString();//probably different - wants type of currency, not value in local currency
                            ExtractRow["MEINS - Sales Unit"] = "";//?
                            ExtractRow["ABSMG - Sales quantity"] = inputstring[74].ToString();
                            ExtractRow[inputstring[6].ToString()] = inverseString(inputstring[76].ToString());
                            ExtractRow["VV230 - Sales Volume KG"] = inputstring[78].ToString();
                            ExtractRow["VV998 - Periodic Quantity SQM"] = inputstring[77].ToString();
                            ExtractRow["0FISCYEAR-Planning Year"] = inputstring[73].ToString();
                            ExtractRow["WWPER - Fiscal Period"] = inputstring[71].ToString();

                            dtx.Rows.Add(ExtractRow);

                        }

                        else
                            ExtractRow[inputstring[6].ToString()] = inverseString(inputstring[76].ToString());
                        //stop cut and paste
                        //set output itemprev for comparision
                        outputItemPrev = outputItem;
                        versionPrev = version;
                        counter++;
                        prevKey = inputstring[0].ToString();
                        
                    }
                }
                // fileHeaderRows.ToCSV("C:/output/filenamestransactional.csv");
                dtx.ToCSV("C:/output/"+ path.ToString() + ".csv");
                dtx.Clear();
            }
        }


        public static DataTable OutputDataTable1()
        {
            DataTable dt = new DataTable();


            DataColumn OutputColumn;

            //BUDAT - Posting Date 1
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "BUDAT - Posting Date";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //FDAT - Invoice Date
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "FDAT - Invoice Date";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //BURKS - Company Code
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "BURKS - Company Code";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //VERSI - Version
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VERSI - Version";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CPLYEAR - Planning Year 5
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CPLYEAR - Planning Year";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //0FISCYEAR-Planning Year
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "0FISCYEAR-Planning Year";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWPER - Fiscal Period
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPER - Fiscal Period";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CCOMPANY - Company
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CCOMPANY - Company";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //VBUND - Partner Company
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VBUND - Partner Company";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CFCH002620 - Item
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CFCH002620 - Item";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KOKRS - Controlling Area -11
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KOKRS - Controlling Area";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //PRCTR - Profit Center
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "PRCTR - Profit Center";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);


            //CPPRCTR - Partner Profit Center
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CPPRCTR - Partner Profit Center";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWBRN - Branch/Industry
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWBRN - Branch/Industry";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWPRG - Product Group
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPRG - Product Group";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);


            //WWPPG - Partner Product Group
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPPG - Partner Product Group";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWART - Material Number
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWART - Material Number";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KUNWE - Ship-To (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KUNWE - Ship-To (local)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KNDNR - Sold-To (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KNDNR - Sold-To (local)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KUNRE - Bill-To (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KUNRE - Bill-To (local)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KUNRG - Payer (local) - 21
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KUNRG - Payer (local)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWKUN - Ship-To Final (local)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWKUN - Ship-To Final (local)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWLWE - Country (Ship-To)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWLWE - Country (Ship-To)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //LAND1 - Country (Sold-To)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "LAND1 - Country (Sold-To)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWLRE - Country (Bill-To) 25
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWLRE - Country (Bill-To)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWLRG - Country (Payer)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWLRG - Country (Payer)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWFCU - Country (Ship-To Final)
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWFCU - Country (Ship-To Final)";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KSTRG - Cost object
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KSTRG - Cost object";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWKAF - Sales order
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWKAF - Sales order";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //KDPOS - Sales order item
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "KDPOS - Sales order item";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWREN - CF invoice number - 31
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWREN - CF invoice number";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);


            //CFCH00056 - Bill. Item
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CFCH00056 - Bill. Item";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWBUN - BU
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWBUN - BU";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWPST - Product Structure
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPST - Product Structure";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWIDS - Identstring
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWIDS - Identstring";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWPRS - Product segment
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWPRS - Product segment";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWHWK - Product characteristic
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWHWK - Product characteristic";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWFTT - Lacquering
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWFTT - Lacquering";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWKAS- Coating
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWKAS- Coating";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWDRU - Print
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWDRU - Print";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWEND - Final Form - 41
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWEND - Final Form";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //WWBRA - Brand
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "WWBRA - Brand";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //MATKL - Material Group
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "MATKL - Material Group";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //FRWAE - Local Currency
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "FRWAE - Local Currency";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10010ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10010ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10015ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10015ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10020ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10020ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10040ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10040ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10050ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10050ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10070ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10070ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10080ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10080ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10090ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10090ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10100ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10100ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10130ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10130ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10140ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10140ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10150ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10150ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10170ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10170ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10180ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10180ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10190ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10190ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10200ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10200ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10210ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10210ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10250ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10250ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10260ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10260ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10270ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10270ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10280ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10280ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10290ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10290ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10300ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10300ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10330ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10330ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10340ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10340ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10350ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10350ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10360ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10360ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10370ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10370ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10380ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10380ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10410ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10410ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10420ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10420ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10430ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10430ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10440ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10440ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10450ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10450ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10460ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10460ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10490ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10490ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10500ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10500ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10510ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10510ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10520ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10520ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10530ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10530ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10540ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10540ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10550ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10550ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10560ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10560ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10570ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10570ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10580ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10580ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10610ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10610ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10620ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10620ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10630ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10630ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10650ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10650ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10660ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10660ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10670ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10670ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10680ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10680ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10700ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10700ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10710ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10710ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10720ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10720ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10730ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10730ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10740ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10740ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10760ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10760ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10770ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10770ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10810ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10810ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10820ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10820ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10830ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10830ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10840ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10840ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10850ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10850ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10860ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10860ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10870ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10870ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10910ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10910ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10920ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10920ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10930ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10930ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10940ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10940ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10950ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10950ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10960ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10960ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10980ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10980ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM10990ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM10990ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11000ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11000ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11020ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11020ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11030ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11030ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11040ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11040ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11050SUM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11050SUM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11070ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11070ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11080ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11080ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11090ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11090ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11100ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11100ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11120ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11120ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11125ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11125ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11140ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11140ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11145ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11145ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //CM11160ITM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "CM11160ITM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //MEINS - Sales Unit
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "MEINS - Sales Unit";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //ABSMG - Sales quantity
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "ABSMG - Sales quantity";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //VV230 - Sales Volume KG
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VV230 - Sales Volume KG";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            //VV998 - Periodic Quantity SQM
            OutputColumn = new DataColumn();
            OutputColumn.DataType = System.Type.GetType("System.String");
            OutputColumn.ColumnName = "VV998 - Periodic Quantity SQM";
            OutputColumn.DefaultValue = "0";
            dt.Columns.Add(OutputColumn);

            return dt;





        }

        public static int convertItem(string itemInput)
        {

            return Convert.ToInt16(itemInput.Substring(2, 5));
            

        }

        public static string convertDate(string lineDate)
        {

            return lineDate.Substring(4, 2)+"/"+lineDate.Substring(6, 2)+"/"+lineDate.Substring(0, 4);
        }

        public static string convertDate(string month, string year)
        {
            /*if (month.Substring(1, 1) == "0")
                month = month.Substring(2, 1);
            else
                month = month.Substring(1, 2); */

         

            return month + "/" + DateTime.DaysInMonth(Convert.ToInt16(year), Convert.ToInt16(month)).ToString() + "/" + year;
        }

        

        /* public static string convertDateDot(string lineDate)
         {
             return lineDate.Substring()

         } */

        public static string inverseString(string input)
        {
            double i = 0;
            input = input.Replace("\"", "");

            if(input.Substring(0,1)=="-")
            {
                i = Convert.ToDouble(input) * (-1);
                return i.ToString();
            }
            if (input.Contains("-"))
            {
                

                i = Convert.ToDouble(input.Substring(0, input.Length - 1));
                return i.ToString();
            } 

            i = Convert.ToDouble(input)*(-1);
            return i.ToString();
            
        }

        //multi-file load for stream process
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            //Stream myStream = null;
            openFileDialog1.InitialDirectory = "C:\\Users'\'john.mark.kennedy'\'Documents'\'";
            //openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }


            }

            //DataTable outputtable = Form1.OutputDataTable1();
            //loop through selection populating global string array varaible

            DataRow fileArrayRow;

            foreach (String file in openFileDialog1.FileNames)
            {

                fileArrayRow = fileArray.NewRow();
                fileArrayRow[0] = file;
                fileArray.Rows.Add(fileArrayRow);
                //MessageBox.Show(fileArray.Rows[0][0].ToString());
            }

            //stop here
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int counter = 0;
            StreamReader streamreaderTXT;
            char[] delimiter = new char[] { ',', ';', '\t' };
            string path = "";
            string outfile = "c:\\output\\";
            StreamWriter streamWriterTXT;
            DataRow ExtractRow;

            using (DataTable dtx = OutputDataTable1())
            {
                foreach (DataRow dt in fileArray.Rows)
                {
                    counter++;
                    streamreaderTXT = new StreamReader(dt[0].ToString());
                    string[] inputstring = streamreaderTXT.ReadLine().Split(delimiter);
                    path = Path.GetFileNameWithoutExtension(dt[0].ToString());
                    string prevKey = "";
                    ExtractRow = dtx.NewRow();

                    // MessageBox.Show(outfile + path);
                    streamWriterTXT = new StreamWriter(outfile + path + ".csv");
                    //streamWriterTXT.WriteLine(String.Join(";", inputstring));
                    //streamWriterTXT.WriteLine("here");


                    while (!streamreaderTXT.EndOfStream)
                    {
                        inputstring = streamreaderTXT.ReadLine().Split(delimiter);
                       // MessageBox
                        //a key was added to align items - prevKey is the prior key value
                        //if prevkey =="" signifies this is the beginning of the file
                        //if the current key (position 0) is not equal to the previous key, this represents that the item is different and
                        // a new row should be started.
                        if (prevKey == "" || inputstring[0].ToString() != prevKey)
                        {
                            if (prevKey != "")
                            {
                                streamWriterTXT.WriteLine(String.Join(",", ExtractRow.ItemArray));
                                dtx.Clear();
                                ExtractRow = dtx.NewRow();

                               
                            }

                            //adding to individual columns CFC072

                            {
                                ExtractRow["BURKS - Company Code"] = inputstring[9].ToString();  //need to figure this piece out?
                                ExtractRow["VERSI - Version"] = inputstring[3].ToString();
                                ExtractRow["CPLYEAR - Planning Year"] = inputstring[4].ToString();

                                ExtractRow["CCOMPANY - Company"] = inputstring[8].ToString();
                                //not available
                                ExtractRow["VBUND - Partner Company"] = "";
                                ExtractRow["CFCH002620 - Item"] = "";
                                ExtractRow["KOKRS - Controlling Area"] = inputstring[10].ToString();
                                ExtractRow["PRCTR - Profit Center"] = inputstring[13].ToString();
                                ExtractRow["CPPRCTR - Partner Profit Center"] = inputstring[11].ToString();
                                ExtractRow["WWBRN - Branch/Industry"] = inputstring[15].ToString();
                                ExtractRow["WWBUN - BU"] = inputstring[16].ToString();
                                ExtractRow["WWPRG - Product Group"] = inputstring[36].ToString();
                                ExtractRow["WWPPG - Partner Product Group"] = inputstring[37].ToString();
                                ExtractRow["WWART - Material Number"] = inputstring[46].ToString();//
                                ExtractRow["KUNWE - Ship-To (local)"] = inputstring[20].ToString();
                                ExtractRow["KNDNR - Sold-To (local)"] = inputstring[28].ToString();
                                ExtractRow["KUNRE - Bill-To (local)"] = inputstring[32].ToString();
                                ExtractRow["KUNRG - Payer (local)"] = inputstring[34].ToString();
                                ExtractRow["WWKUN - Ship-To Final (local)"] = inputstring[24].ToString();
                                ExtractRow["WWLWE - Country (Ship-To)"] = inputstring[21].ToString();
                                ExtractRow["LAND1 - Country (Sold-To)"] = inputstring[29].ToString();
                                ExtractRow["WWLRE - Country (Bill-To)"] = inputstring[33].ToString();
                                ExtractRow["WWLRG - Country (Payer)"] = inputstring[35].ToString();
                                ExtractRow["WWFCU - Country (Ship-To Final)"] = inputstring[25].ToString();

                                ExtractRow["KSTRG - Cost object"] = ""; //missing
                                ExtractRow["WWKAF - Sales order"] = inputstring[57].ToString();
                                ExtractRow["KDPOS - Sales order item"] = inputstring[52].ToString();
                                ExtractRow["WWREN - CF invoice number"] = inputstring[56].ToString();
                                ExtractRow["CFCH00056 - Bill. Item"] = ""; //missing

                                ExtractRow["WWPST - Product Structure"] = inputstring[45].ToString();
                                ExtractRow["WWIDS - Identstring"] = inputstring[44].ToString();
                                ExtractRow["WWPRS - Product segment"] = inputstring[38].ToString();
                                ExtractRow["WWHWK - Product characteristic"] = inputstring[39].ToString();
                                ExtractRow["WWFTT - Lacquering"] = inputstring[40].ToString();
                                ExtractRow["WWKAS- Coating"] = inputstring[41].ToString();
                                ExtractRow["WWDRU - Print"] = inputstring[42].ToString();
                                ExtractRow["WWEND - Final Form"] = inputstring[43].ToString();
                                ExtractRow["WWBRA - Brand"] = inputstring[49].ToString();
                                ExtractRow["MATKL - Material Group"] = inputstring[48].ToString();
                                ExtractRow["BUDAT - Posting Date"] = convertDate(inputstring[67].ToString(), inputstring[69].ToString());
                                ExtractRow["FDAT - Invoice Date"] = convertDate(inputstring[67].ToString(), inputstring[69].ToString());
                                ExtractRow["FRWAE - Local Currency"] = inputstring[0].ToString();
                                ExtractRow["MEINS - Sales Unit"] = "";//?
                                ExtractRow["ABSMG - Sales quantity"] = inputstring[70].ToString();
                                ExtractRow[inputstring[6].ToString()] = inverseString(inputstring[72].ToString());
                                ExtractRow["VV230 - Sales Volume KG"] = inputstring[75].ToString();
                                ExtractRow["VV998 - Periodic Quantity SQM"] = inputstring[73].ToString();
                                ExtractRow["0FISCYEAR-Planning Year"] = inputstring[69].ToString();
                                ExtractRow["WWPER - Fiscal Period"] = inputstring[67].ToString();

                                dtx.Rows.Add(ExtractRow);
                            }
                        }

                        else
                            ExtractRow[inputstring[6].ToString()] = inverseString(inputstring[72].ToString());

                        counter++;
                        prevKey = inputstring[0].ToString();

                       

                    }
                    streamWriterTXT.WriteLine(String.Join(",", ExtractRow.ItemArray));
                    streamWriterTXT.Close();

                    MessageBox.Show("done");
                }
            }
        }

      /*  public static string returnString(DataTable dt)
        {
            string String1 = "";
            foreach (DataColumn dc in dt.Columns)
            {
                String1+=dc
            }
            return "";
        } */
    }

   


}



