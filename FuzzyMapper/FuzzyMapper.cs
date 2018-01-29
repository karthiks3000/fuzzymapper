using Infragistics.Documents.Excel;
using Infragistics.Win;
using Infragistics.Win.UltraWinEditors;
using Infragistics.Win.UltraWinGrid;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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

        /// <summary>
        /// Reads the provided excel & sheet and returns a DataTable. If no sheet name is provided then the first sheet of the excel is chosen.
        /// </summary>
        /// <param name="excelname"></param>
        /// <param name="sheetname"></param>
        /// <returns>DataTable</returns>
        private DataTable LoadExcel(string excelname, string sheetname = "")
        {
            var excelDataTable = new DataTable();
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
                    this.upbProgressBar.Value = 0;
                    this.upbProgressBar.Maximum = theWorksheet.Rows.Count();
                    this.upbProgressBar.Step = 1;

                    //Iterate through all Worksheet rows
                    foreach (WorksheetRow theWorksheetRow in theWorksheet.Rows)
                    {
                        this.upbProgressBar.PerformStep();
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

                                    theDataColumn.DataType = Type.GetType("System.String");
                                    //theWorksheet.Rows[theRowCounter + 1].Cells[theCellCounter].Value.GetType();
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
            catch (Exception)
            {
                throw;
            }
            finally
            {
                this.Cursor = Cursors.Default;
                this.upbProgressBar.Value = 0;
            }
            return excelDataTable;
        }

        /// <summary>
        /// Update's the specified grid row cells with the values from the provided DataTable using the specified key column name and value
        /// </summary>
        /// <param name="row"></param>
        /// <param name="dt"></param>
        /// <param name="key"></param>
        /// <param name="keyCol"></param>
        /// <param name="colPrefix"></param>
        private void UpdateResultRow(UltraGridRow row, DataTable dt, string key, string keyCol, string colPrefix = "")
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
                        row.Cells[colPrefix + col.ColumnName].Value = dbl;
                    }
                    else
                        row.Cells[colPrefix + col.ColumnName].Value = DestinationRow[col.ColumnName].ToString();
                }                
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Updates the grid header with the total row count and the visible row count
        /// </summary>
        /// <param name="uGrid"></param>
        private void UpdateGridRowCounts(UltraGrid uGrid)
        {
            try
            {
                uGrid.Text = (uGrid.Rows.Count).ToString() + " Row(s) / " + (uGrid.Rows.VisibleRowCount - 1).ToString() + " Visible Row(s)";
            }
            catch (Exception)
            {
                throw;
            }
        }


        private bool LoadGrid(UltraGrid uGrid,ref DataTable dt,string excelName,string columnType,UltraComboEditor uCombo)
        {
            try
            {
                uGrid.DataSource = null;
                dt = LoadExcel(excelName);
                uGrid.DataSource = dt;
                ValueList colValueList = new ValueList();

                uCombo.Items.Clear();

                foreach (DataColumn col in dt.Columns)
                {
                    colValueList.ValueListItems.Add(col.ColumnName, col.ColumnName);
                    uCombo.Items.Add(col.ColumnName, col.ColumnName);
                }
                this.udsMapColumns.Rows.Clear();

                this.ugMapColumns.DisplayLayout.Bands[0].Columns[columnType].ValueList = colValueList;
                this.ugMapColumns.DisplayLayout.Bands[0].Columns[columnType].Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate;
                this.ugMapColumns.DisplayLayout.Bands[0].Columns[columnType].AutoCompleteMode = Infragistics.Win.AutoCompleteMode.SuggestAppend;

                uCombo.SelectedIndex = 0;
                UpdateGridRowCounts(uGrid);
            }
            catch (Exception)
            {
                throw;                
            }
            return true;
        }
        #endregion Private Methods

        #region Events

        private void ubtnSource_Click(object sender, EventArgs e)
        {
            try
            {
                this.openFileDialog1.ShowDialog();
                this.utxtSource.Text = this.openFileDialog1.FileName;

                if (!this.utxtSource.Text.Equals(""))
                {
                    this.Cursor = Cursors.WaitCursor;
                    LoadGrid(this.ugSource, ref SourceDataTable, this.utxtSource.Text, "SourceColumn", this.ucSourceKeyCol);
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
                    LoadGrid(this.ugDestination, ref DestinationDataTable, this.utxtDestination.Text, "DestinationColumn", this.ucDestinationKeyCol);
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

                if (this.SourceDataTable != null && this.DestinationDataTable != null)
                {
                    this.Cursor = Cursors.WaitCursor;
                    float fuzzyness = float.Parse(umeAccuracy.Text) / 100;

                    this.upbProgressBar.Value = 0;
                    this.upbProgressBar.Minimum = 0;
                    Dictionary<string, string> DestinationValues = default(Dictionary<string, string>);
                    Dictionary<int, Dictionary<string, string>> DestinationValueLists = new Dictionary<int, Dictionary<string, string>>();

                    foreach (UltraGridRow urow in this.ugMapColumns.Rows)
                    {
                        DestinationValues = this.DestinationDataTable.AsEnumerable().Select(row => new
                        {
                            attribute1_name = row[this.ucDestinationKeyCol.SelectedItem.DisplayText].ToString(),
                            attribute2_name = row[urow.Cells["DestinationColumn"].Value.ToString()].ToString()
                        }).Distinct().ToDictionary(s => s.attribute1_name, s => s.attribute2_name);

                        DestinationValueLists.Add(urow.Index, DestinationValues);
                    }

                    if (this.ucMapType.SelectedItem.DisplayText.Equals("Map"))
                    {
                        this.ugResults.DataSource = null;
                        this.ugResults.DataSource = this.SourceDataTable.Copy();
                        foreach (DataColumn col in this.DestinationDataTable.Columns)
                        {
                            this.ugResults.DisplayLayout.Bands[0].Columns.Add("Des-" + col.ColumnName);
                        }

                        this.upbProgressBar.Step = 1;
                        this.upbProgressBar.Maximum = this.ugResults.Rows.Count;

                        foreach (UltraGridRow row in this.ugResults.Rows)
                        {
                            this.upbProgressBar.PerformStep();
                            Dictionary<int, Dictionary<string, string>> FoundMatchList = new Dictionary<int, Dictionary<string, string>>();
                            Dictionary<string, string> FoundMatches = new Dictionary<string, string>();

                            bool matchFound = true;
                            try
                            {
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
                                    FoundMatchList.Add(urow.Index, FoundMatches);
                                }
                            }
                            catch (Exception)
                            {
                                throw;
                            }

                            if (matchFound)
                            {
                                try
                                {
                                    Dictionary<string, string> ExactMatches = new Dictionary<string, string>();
                                    if (FoundMatchList.Count == 1)
                                    {
                                        ExactMatches = FoundMatchList[0];
                                    }
                                    else
                                    {
                                        var intersectValues = FoundMatchList[0].Keys.ToList();
                                        for (int i = 0; i < FoundMatchList.Count - 1; i++)
                                        {
                                            //var current = new Dictionary<string, string>(FoundMatchList[i]);
                                            var next = new Dictionary<string, string>(FoundMatchList[i + 1]);
                                            intersectValues = next.Keys.Where(x => intersectValues.Contains(x)).ToList();
                                        }
                                        if (intersectValues.Count > 0)
                                            ExactMatches.Add(intersectValues.FirstOrDefault(), "");
                                    }
                                    if (ExactMatches.Count > 0)
                                        UpdateResultRow(row, this.DestinationDataTable, ExactMatches.FirstOrDefault().Key, this.ucDestinationKeyCol.SelectedItem.DisplayText, "Des-");
                                }
                                catch (Exception)
                                {
                                    throw;
                                }
                            }
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

                        this.upbProgressBar.Step = 1;
                        this.upbProgressBar.Maximum = this.ugSource.Rows.Count;

                        foreach (UltraGridRow row in this.ugSource.Rows)
                        {
                            this.upbProgressBar.PerformStep();
                            Dictionary<string, string> FoundMatches = new Dictionary<string, string>();
                            Dictionary<int, Dictionary<string, string>> FoundMatchList = new Dictionary<int, Dictionary<string, string>>();

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

                                FoundMatchList.Add(urow.Index, FoundMatches);
                            }
                            if (matchFound)
                            {
                                Dictionary<string, string> ExactMatches = new Dictionary<string, string>();
                                if (FoundMatchList.Count == 1)
                                {
                                    ExactMatches = FoundMatchList[0];
                                }
                                else
                                {
                                    var intersectValues = FoundMatchList[0].Keys.ToList();
                                    for (int i = 0; i < FoundMatchList.Count - 1; i++)
                                    {
                                        //var current = new Dictionary<string, string>(FoundMatchList[i]);
                                        var next = new Dictionary<string, string>(FoundMatchList[i + 1]);
                                        intersectValues = next.Keys.Where(x => intersectValues.Contains(x)).ToList();
                                    }
                                    foreach (var item in intersectValues)
                                    {
                                        if (!ExactMatches.ContainsKey(item))
                                            ExactMatches.Add(item, "");
                                    }
                                }

                                foreach (var item in ExactMatches)
                                {
                                    DataRow dr = dtResults.NewRow();
                                    dtResults.Rows.Add(dr);

                                    UpdateResultRow(this.ugResults.Rows[dtResults.Rows.IndexOf(dr)], this.DestinationDataTable, item.Key, this.ucDestinationKeyCol.SelectedItem.DisplayText, "Des-");
                                    UpdateResultRow(this.ugResults.Rows[dtResults.Rows.IndexOf(dr)], this.SourceDataTable, row.Cells[this.ucSourceKeyCol.SelectedItem.DisplayText].Value.ToString(), this.ucSourceKeyCol.SelectedItem.DisplayText);
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
                UpdateGridRowCounts(this.ugResults);

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

                    if (this.utcTabControl.ActiveTab == this.ultraTabPageControl1.Tab)
                        this.ultraGridExcelExporter1.Export(this.ugSource, savefile.FileName);
                    else if (this.utcTabControl.ActiveTab == this.ultraTabPageControl2.Tab)
                        this.ultraGridExcelExporter1.Export(this.ugDestination, savefile.FileName);
                    else if (this.utcTabControl.ActiveTab == this.ultraTabPageControl3.Tab)
                        this.ultraGridExcelExporter1.Export(this.ugResults, savefile.FileName);

                    savefile.OpenFile();
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

        #endregion Events

        private void ugSource_AfterRowFilterChanged(object sender, AfterRowFilterChangedEventArgs e)
        {
            UpdateGridRowCounts(ugSource);
        }

        private void ugDestination_AfterRowFilterChanged(object sender, AfterRowFilterChangedEventArgs e)
        {
            UpdateGridRowCounts(ugDestination);
        }

        private void ugResults_AfterRowFilterChanged(object sender, AfterRowFilterChangedEventArgs e)
        {
            UpdateGridRowCounts(ugResults);
        }
    }
}