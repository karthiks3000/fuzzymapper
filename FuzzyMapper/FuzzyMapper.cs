using Infragistics.Documents.Excel;
using Infragistics.Win;
using Infragistics.Win.UltraWinEditors;
using Infragistics.Win.UltraWinGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FuzzyMapper
{
    public partial class FuzzyMapper : Form
    {
        DataTable SourceDataTable;
        DataTable DestinationDataTable;

        public FuzzyMapper()
        {
            InitializeComponent();
            this.ucMapType.SelectedIndex = 0;
            this.ucAlgorithm.SelectedIndex = 0;
        }

        #region Private Methods
        
        private DataTable LoadExcel(string excelname, string sheetname = "")
        {
            DataTable excelDataTable = new DataTable();
            try
            {
                this.Cursor = Cursors.WaitCursor;
                if (File.Exists(excelname))
                {
                    //Load the Excel File into the Workbook Object
                    Workbook theWorkbook = Workbook.Load(excelname);
                    Worksheet theWorksheet;
                    if (sheetname.Equals(""))
                    {
                        theWorksheet = theWorkbook.Worksheets[0];
                    }
                    else
                    {
                        theWorksheet = theWorkbook.Worksheets[sheetname];
                    }
                    
                    //We will place the Excel Data into this DataTable                    

                    int theRowCounter = 0;
                    int theCellCounter = 0;

                    //Iterate through all Worksheet rows
                    foreach (WorksheetRow theWorksheetRow in theWorksheet.Rows)
                    {
                        if (theRowCounter == 0)
                        {
                            //This is the Header Row. We are assuming that the Excel Worksheet's 
                            //first row contains the schema of our soon to be data model. 
                            //We will use this information to build our DataTable's schema
                            foreach (WorksheetCell theWorksheetCell in theWorksheetRow.Cells)
                            {

                                string theCellValue = theWorksheetCell.Value.ToString().Trim();

                                if (theCellValue != string.Empty)
                                {
                                    //This is the "Header Row"
                                    //Create a DataColumn for each Column taken from the first Worksheet row
                                    DataColumn theDataColumn = excelDataTable.Columns.Add();

                                    //Since this is the Header Row, we use the cell value
                                    //as the Column Name
                                    theDataColumn.ColumnName = theCellValue;

                                    //Here we skip to the actual data row (the row below the header row) and
                                    //set the data column's data type to the type that exists in the corresponding
                                    //cell in the actual data row of the Worksheet
                                    theDataColumn.DataType =
                                        theWorksheet.Rows[theRowCounter + 1].Cells[theCellCounter].Value.GetType();
                                }
                                else
                                {
                                    break;
                                    //Exit the loop so that we do not
                                    //traverse all empty Worksheet Cells.
                                }

                                theCellCounter++;
                            }
                        }
                        else
                        //This is the actual data that will populate the data model
                        {
                            theCellCounter = 0;

                            //Add a new empty data row to our data model        
                            DataRow theDataRow = excelDataTable.NewRow();

                            //iterate through each current Worksheet cell and populate the new data row.
                            foreach (WorksheetCell theWorksheetCell in theWorksheetRow.Cells)
                            {
                                object theValue = theWorksheet.Rows[theRowCounter].Cells[theCellCounter].Value;

                                if (theValue != null)
                                {
                                    theDataRow[theCellCounter] = theValue;
                                }
                                else
                                {
                                    break;
                                    //Exit the loop so that we do not
                                    //traverse all empty Worksheet Cells.
                                }

                                theCellCounter++;
                            }

                            //Add the Data Row to the DataTable
                            excelDataTable.Rows.Add(theDataRow);
                        }

                        theRowCounter++;
                    }

                    //AcceptChanges so that these do not appear to be NEW rows
                    excelDataTable.AcceptChanges();
                }

            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
            return excelDataTable;

        }

        private void AddResultRow(UltraGridRow row, DataTable dt, string key, string keyCol, string colprefix = "")
        {
            try
            {
                DataRow DestinationRow = dt.AsEnumerable().Where(d => d[keyCol].ToString().Equals(key)).FirstOrDefault();

                foreach (DataColumn col in dt.Columns)
                {
                    if (col.DataType == Type.GetType("System.Double"))
                    {
                        Double dbl = 0;
                        Double.TryParse(DestinationRow[col.ColumnName].ToString(), out dbl);
                        row.Cells[colprefix + col.ColumnName].Value = dbl;

                    }
                    else
                        row.Cells[colprefix + col.ColumnName].Value = DestinationRow[col.ColumnName].ToString();

                }
                row.Cells["Map%"].Value = umeAccuracy.Text;

            }
            catch (Exception e)
            {

                throw e;
            }
        }

        #endregion
        private void ubtnSource_Click(object sender, EventArgs e)
        {
            try
            {
                this.openFileDialog1.ShowDialog();
                this.utxtSource.Text = this.openFileDialog1.FileName;

                if (!this.utxtSource.Text.Equals(""))
                {
                    this.Cursor = Cursors.WaitCursor;

                    this.ugSource.DataSource = null;
                    SourceDataTable = LoadExcel(this.utxtSource.Text);
                    this.ugSource.DataSource = SourceDataTable;
                    ValueList ucSourceCol = new ValueList();
                    //this.ucSourceCol.Items.Clear();
                    this.ucSourceKeyCol.Items.Clear();

                    foreach (DataColumn col in SourceDataTable.Columns)
                    {
                        ucSourceCol.ValueListItems.Add(col.ColumnName, col.ColumnName);
                        //this.ucSourceCol.Items.Add(col.ColumnName, col.ColumnName);
                        this.ucSourceKeyCol.Items.Add(col.ColumnName, col.ColumnName);
                    }
                    this.ugMapColumns.DisplayLayout.Bands[0].Columns["SourceColumn"].ValueList = ucSourceCol;
                    //this.ucSourceCol.SelectedIndex = 0;
                    this.ucSourceKeyCol.SelectedIndex = 0;
                    this.utcTabControl.SelectedTab = this.ultraTabPageControl1.Tab;

                }
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }
      
        private void ubtnDestination_Click(object sender, EventArgs e)
        {
            try
            {
                this.openFileDialog2.ShowDialog();
                this.utxtDestination.Text = this.openFileDialog2.FileName;
                if (!this.utxtDestination.Text.Equals(""))
                {
                    this.Cursor = Cursors.WaitCursor;

                    DestinationDataTable = LoadExcel(this.utxtDestination.Text);
                    this.ugDestination.DataSource = DestinationDataTable;
                    //this.ucDestinationCol.Items.Clear();
                    this.ucDestinationKeyCol.Items.Clear();
                    ValueList ucDestinationCol = new ValueList();

                    foreach (DataColumn col in DestinationDataTable.Columns)
                    {
                        ucDestinationCol.ValueListItems.Add(col.ColumnName, col.ColumnName);
                        //this.ucDestinationCol.Items.Add(col.ColumnName, col.ColumnName);
                        this.ucDestinationKeyCol.Items.Add(col.ColumnName, col.ColumnName);
                    }
                    this.ugMapColumns.DisplayLayout.Bands[0].Columns["DestinationColumn"].ValueList = ucDestinationCol;

                    //this.ucDestinationCol.SelectedIndex = 0;
                    this.ucDestinationKeyCol.SelectedIndex = 0;
                    this.utcTabControl.SelectedTab = this.ultraTabPageControl2.Tab;

                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                    this.Cursor = Cursors.Default;
            }
        }                     

        private void ubtnMap_Click(object sender, EventArgs e)
        {
            try
            {
                string source = "";
                string sourceColumnName = "";//this.ucSourceCol.SelectedItem.DisplayText;


                if (this.SourceDataTable != null && this.DestinationDataTable != null)
                {
                    this.Cursor = Cursors.WaitCursor;
                    float fuzzyness = float.Parse(umeAccuracy.Text) / 100;

                    this.upbProgressBar.Value = 0;
                    this.upbProgressBar.Minimum = 0;
                    Dictionary<string, string> DestinationValues = default(Dictionary<string, string>);
                    Dictionary<int, Dictionary<string,string>> DestinationValueLists = new Dictionary<int, Dictionary<string, string>>();

                    foreach (UltraGridRow urow in this.ugMapColumns.Rows)
                    {

                        DestinationValues = this.DestinationDataTable.AsEnumerable().Select(row => new
                        {
                            attribute1_name = row[this.ucDestinationKeyCol.SelectedItem.DisplayText].ToString(),
                            attribute2_name = row[urow.Cells["DestinationColumn"].Value.ToString()].ToString()
                        }).Distinct().ToDictionary(s => s.attribute1_name, s => s.attribute2_name);


                        DestinationValueLists.Add(urow.Index, DestinationValues);
                    }
                    
                    
                    if (this.ucMapType.SelectedIndex == 0)
                    {
                        this.ugResults.DataSource = null;
                        this.ugResults.DataSource = this.SourceDataTable.Copy();
                        foreach (DataColumn col in this.DestinationDataTable.Columns)
                        {
                            if (this.ugResults.DisplayLayout.Bands[0].Columns.IndexOf("Des-" + col.ColumnName) < 0)
                                this.ugResults.DisplayLayout.Bands[0].Columns.Add("Des-" + col.ColumnName);
                        }
                        this.ugResults.DisplayLayout.Bands[0].Columns.Add("Map%");

                        Dictionary<string, string> FoundMatches = new Dictionary<string, string>();

                        this.upbProgressBar.Step = 1;
                        this.upbProgressBar.Maximum= this.ugResults.Rows.Count;

                        foreach (UltraGridRow row in this.ugResults.Rows)
                        {
                            this.upbProgressBar.PerformStep();
                            bool matchFound = true;
                            foreach (UltraGridRow urow in this.ugMapColumns.Rows)
                            {
                                source = row.Cells[urow.Cells["SourceColumn"].Value.ToString()].Value.ToString();
                                FoundMatches = null;
                                FoundMatches = FuzzySearch.Search_v3(source, DestinationValueLists[urow.Index], fuzzyness, this.ucAlgorithm.SelectedItem.DisplayText);
                                if (FoundMatches.Count == 0)
                                {
                                    matchFound = false;
                                    break;
                                }
                                //FoundMatches = null;
                            }

                            if(matchFound)
                                AddResultRow(row, this.DestinationDataTable, FoundMatches.FirstOrDefault().Key, this.ucDestinationKeyCol.SelectedItem.DisplayText, "Des-");                           
                        }
                    }
                    else
                    {
                        this.ugResults.DataSource = null;
                        DataTable dtResults = this.SourceDataTable.Clone();
                        this.ugResults.DataSource = dtResults;
                        foreach (DataColumn col in this.DestinationDataTable.Columns)
                        {
                            if (this.ugResults.DisplayLayout.Bands[0].Columns.IndexOf("Des-" + col.ColumnName) < 0)
                                this.ugResults.DisplayLayout.Bands[0].Columns.Add("Des-" + col.ColumnName);
                        }
                        this.ugResults.DisplayLayout.Bands[0].Columns.Add("Map%");

                        Dictionary<string, string> FoundMatches = new Dictionary<string, string>();
                        this.upbProgressBar.Step = 1;
                        this.upbProgressBar.Maximum = this.ugSource.Rows.Count;

                        foreach (UltraGridRow row in this.ugSource.Rows)
                        {
                            this.upbProgressBar.PerformStep();
                            bool matchFound = true;
                            foreach (UltraGridRow urow in this.ugMapColumns.Rows)
                            {
                                source = row.Cells[urow.Cells["SourceColumn"].Value.ToString()].Value.ToString();
                                FoundMatches = null;
                                FoundMatches = FuzzySearch.Search_v3(source, DestinationValueLists[urow.Index], fuzzyness, this.ucAlgorithm.SelectedItem.DisplayText);
                                if (FoundMatches.Count == 0)
                                {
                                    matchFound = false;
                                    break;
                                }
                            }
                            if (matchFound)
                            {
                                foreach (var item in FoundMatches)
                                {
                                    DataRow dr = dtResults.NewRow();
                                    dtResults.Rows.Add(dr);

                                    AddResultRow(this.ugResults.Rows[dtResults.Rows.IndexOf(dr)], this.DestinationDataTable, item.Key, this.ucDestinationKeyCol.SelectedItem.DisplayText, "Des-");
                                    AddResultRow(this.ugResults.Rows[dtResults.Rows.IndexOf(dr)], this.SourceDataTable, row.Cells[this.ucSourceKeyCol.SelectedItem.DisplayText].Value.ToString(), this.ucSourceKeyCol.SelectedItem.DisplayText);
                                }
                            }
                        }
                    }
                    this.utcTabControl.SelectedTab = this.ultraTabPageControl3.Tab;

                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.upbProgressBar.Value = 0;
            }
        }
                
        private void ubtnExport_Click(object sender, EventArgs e)
        {
            try
            {
                SaveFileDialog savefile = new SaveFileDialog();
                // set a default file name
                savefile.FileName = "FuzzyMapper.xlsx";
                // set filters - this can be done in properties as well
                savefile.Filter = "Excel files (*.xlsx)|*.xls|All files (*.*)|*.*";

                if (savefile.ShowDialog() == DialogResult.OK)
                {
                    this.Cursor = Cursors.WaitCursor;

                    if(this.utcTabControl.ActiveTab == this.ultraTabPageControl1.Tab)
                        this.ultraGridExcelExporter1.Export(this.ugSource, savefile.FileName);
                    else if (this.utcTabControl.ActiveTab == this.ultraTabPageControl2.Tab)
                        this.ultraGridExcelExporter1.Export(this.ugDestination, savefile.FileName);
                    else if (this.utcTabControl.ActiveTab == this.ultraTabPageControl3.Tab)
                        this.ultraGridExcelExporter1.Export(this.ugResults, savefile.FileName);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void ubtnAdd_Click(object sender, EventArgs e)
        {
            this.udsMapColumns.Rows.Add();
        }

        private void ubtnDelete_Click(object sender, EventArgs e)
        {
            this.ugMapColumns.DeleteSelectedRows();
        }
    }
}
