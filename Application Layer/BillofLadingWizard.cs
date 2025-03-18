using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BillOfLading
{
    internal class BillofLadingWizard
    {
        #region Variables
        SAPbouiCOM.Form oForm;

        SAPbobsCOM.Recordset oRecordSet;
        private SAPbouiCOM.IChooseFromListEvent oCFLEvent;
        private SAPbouiCOM.EditText oEditText;
        private SAPbouiCOM.ComboBox oCombobox;
        private SAPbouiCOM.EditTextColumn oEditTextColumn;
        private SAPbouiCOM.Grid oGrid;



        private SAPbouiCOM.CheckBoxColumn oCheckBox, oCheckBox1;

        public string MatrixId;

        private SAPbouiCOM.Column oColumn;
        private DateTime dtValidFrom, dtValidTo;
        private string strQuery;

        #endregion

        #region LoadForm



        public void LoadblWizard(string aCardCode)
        {
            try
            {
                GlobalVariables.oGFun.LoadXML(oForm, GlobalVariables.BillWizardID, GlobalVariables.BillWizardXML);

                oForm = EventHandler.oApplication.Forms.Item(GlobalVariables.BillWizardID);
                EventHandler.oApplication.Forms.Item(GlobalVariables.BillWizardID).Visible = true;
                EventHandler.oApplication.Forms.Item(GlobalVariables.BillWizardID).Select();
                oForm.Freeze(true);
                //AddChooseFromList(oForm);
                //b.""U_VC"",
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("1000002").Specific;
                //oGrid.DataTable.ExecuteQuery($@"SELECT  'N' AS ""Select"",b.""DocNum"",a.""U_ContNo"", b.""U_VN"" from ""@BLC_BOLD"" b Inner Join ""@BLC_BLD1"" a on b.""DocEntry"" = a.""DocEntry""    WHERE  b.""Canceled""='N' and a.""U_ContNo"" IS NOT NULL and b.""U_VC"" = '{aCardCode}'");
                oGrid.DataTable.ExecuteQuery($@"SELECT  'N' AS ""Select"",b.""DocNum"", b.""U_VN"",b.""U_VC""   FROM ""@BLC_BOLD"" b Inner Join OPCH T1 on TO_NVARCHAR(b.""U_Draft"") = TO_NVARCHAR(T1.""DocEntry"") WHERE T1.""CANCELED""  = 'N' and  T1.""DocStatus"" = 'O' and  T1.""isIns"" = 'Y' and T1.""CardCode"" = '{aCardCode}' and b.""U_Draft"" IS NOT NULL ");

                oEditTextColumn = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("DocNum");
                //oEditTextColumn.ChooseFromListUID = "CFL_6";
                //oEditTextColumn.ChooseFromListAlias = "CardCode";
                oEditTextColumn.LinkedObjectType = "BLC_BOLD";
                oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oGrid.Columns.Item("DocNum").TitleObject.Caption = "BL Number";
                //oGrid.Columns.Item("U_ContNo").TitleObject.Caption = "Container Number";
                oGrid.Columns.Item("U_VC").TitleObject.Caption = "Supplier Code";
                oGrid.Columns.Item("U_VN").TitleObject.Caption = "Supplier Name";
                oGrid.Columns.Item("U_VC").Editable = false;      
                //oGrid.Columns.Item("U_ContNo").Editable = false;
                oGrid.Columns.Item("DocNum").Editable = false;
                oGrid.Columns.Item("U_VN").Editable = false;
                oGrid.Columns.Item("Select").Editable = true;

                oGrid.AutoResizeColumns();
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

                oForm.DataSources.UserDataSources.Add("frm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.DataSources.UserDataSources.Add("To", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.DataSources.UserDataSources.Add("NumAtCard", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.DataSources.UserDataSources.Add("ItmsGrp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.DataSources.UserDataSources.Add("WhsCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.DataSources.UserDataSources.Add("ItemCode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.DataSources.UserDataSources.Add("SubCatNum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                oForm.PaneLevel = 1;
                //oForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        private void AddChooseFromList(SAPbouiCOM.Form aForm)
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = aForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromList oCFL;
                SAPbouiCOM.Conditions oConditions;
                SAPbouiCOM.Condition oCondition;

                // Purchase Order CFL for CFL_4
                oCFL = oCFLs.Item("CFL_4");
                oConditions = (SAPbouiCOM.Conditions)EventHandler.oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                oCondition = oConditions.Add();
                oCondition.BracketOpenNum = 2;
                oCondition.Alias = "DocStatus";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "O";
                oCondition.BracketCloseNum = 1;
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCondition = oConditions.Add();
                oCondition.BracketOpenNum = 1;
                oCondition.Alias = "DocStatus";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "O";
                oCondition.BracketCloseNum = 2;

                oCFL.SetConditions(oConditions);

                // Purchase Order CFL for CFL_3
                oCFL = oCFLs.Item("CFL_3");
                oConditions = (SAPbouiCOM.Conditions)EventHandler.oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                oCondition = oConditions.Add();
                oCondition.BracketOpenNum = 2;
                oCondition.Alias = "DocStatus";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "O";
                oCondition.BracketCloseNum = 1;
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCondition = oConditions.Add();
                oCondition.BracketOpenNum = 1;
                oCondition.Alias = "DocStatus";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "O";
                oCondition.BracketCloseNum = 2;

                oCFL.SetConditions(oConditions);

                // Purchase Order CFL for CFL_6
                oCFL = oCFLs.Item("CFL_6");
                oConditions = (SAPbouiCOM.Conditions)EventHandler.oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);

                oCondition = oConditions.Add();
                oCondition.BracketOpenNum = 1;
                oCondition.Alias = "CardType";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "S";
                oCondition.BracketCloseNum = 1;

                oCFL.SetConditions(oConditions);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        public void DataBind(SAPbouiCOM.Form aform)
        {
            string getIdocnumLists = getDocNum(aform);
            //string containerNo = GetIContainerLists(aform);
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("1").Specific;
            oGrid.DataTable.ExecuteQuery($@"SELECT Distinct a.""U_ContNo"",'N' AS ""Select"",b.""DocNum"", b.""U_VN"",b.""U_VC"" from ""@BLC_BOLD"" b Inner Join ""@BLC_BLD1"" a on b.""DocEntry"" = a.""DocEntry""    WHERE  b.""Canceled""='N' and a.""U_ContNo"" IS NOT NULL and b.""DocNum"" IN ( {getIdocnumLists} )");
            oGrid.Columns.Item("U_VC").TitleObject.Caption = "Supplier Code";
            oEditTextColumn = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("U_VC");
            oEditTextColumn.LinkedObjectType = "2";

            oGrid.Columns.Item("DocNum").TitleObject.Caption = "BL Number";
            oGrid.Columns.Item("U_ContNo").TitleObject.Caption = "Container Number";
            oGrid.Columns.Item("U_VC").TitleObject.Caption = "Supplier Code";
            oGrid.Columns.Item("U_VN").TitleObject.Caption = "Supplier Name";
            oGrid.Columns.Item("U_VC").Editable = false;
            oGrid.Columns.Item("U_ContNo").Editable = false;
            oGrid.Columns.Item("DocNum").Editable = false;
            oGrid.Columns.Item("U_VN").Editable = false;
           

            oGrid.Columns.Item("Select").Editable = true;

            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            oGrid.AutoResizeColumns();
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;
            AssignMatrixLineno(oGrid, aform);
            aform.PaneLevel = 2;
        }

        public void DataBind_ITems(SAPbouiCOM.Form aform)
        {
            string getIdocnumLists = GetItemsLists(aform);
            string containerNo = GetIContainerLists(aform);
            
            //string aCardCode = GetCustomer(aform);
            List<string> DocumentNo = new List<string>();
            string strQuery;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("14").Specific;

      

            strQuery = $@"SELECT 'N' AS ""Select"", 
    T1.""DocEntry"", 
    T1.""DocNum"", 
    T1.""U_VC"", 
    T1.""U_VN"", 
    T0.""LineId"", 
    T0.""U_IC"", 
    T0.""U_IN"", 
    T0.""U_POQ"",
    T0.""U_Commodity"", T0.""U_Model"", T0.""U_ContNo"", T0.""U_ContType"", T0.""U_IncTerms"", 
    T0.""U_FeedPort"", T0.""U_FeedETD"", T0.""U_MainPort"", T0.""U_MainETD"", 
     T0.""U_Pod"", T0.""U_PodETA"", T0.""U_GateDate"", T0.""U_EmptyDate"", 
     T0.""U_CNEE"", T0.""U_Notify"", T0.""U_Forwarder"", T0.""U_ShipLine"", 
     T0.""U_FreeDays"", T0.""U_Cost"", T0.""U_Telex"", T0.""U_TelexDate"",
(Select Distinct F.""PriceBefDi"" from PCH1 F inner join OPCH E on E.""DocEntry"" = F.""DocEntry"" 
     where E.""DocEntry"" = T1.""U_Draft"" and F.""ItemCode"" = T0.""U_IC"") as ""Price""
       FROM ""@BLC_BOLD"" T1
  INNER JOIN ""@BLC_BLD1"" T0 ON T1.""DocEntry"" = T0.""DocEntry""
  WHERE T1.""DocNum"" IN({getIdocnumLists}) and T0.""U_ContNo""
 IN({containerNo})
 and(T0.""U_IC"" is not null or T0.""U_IC"" = '')
     ORDER BY T0.""DocEntry"", T0.""LineId"" ";
            oGrid.DataTable.ExecuteQuery(strQuery);



            oGrid.Columns.Item("U_VC").TitleObject.Caption = "Supplier Code";
            oEditTextColumn = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("U_VC");
            oEditTextColumn.LinkedObjectType = "2";

            oGrid.Columns.Item("U_VN").TitleObject.Caption = "Supplier Name";
            oGrid.Columns.Item("DocNum").TitleObject.Caption = "BL Doc. Number";
            oGrid.Columns.Item("LineId").TitleObject.Caption = "Line No.";
            oGrid.Columns.Item("U_POQ").TitleObject.Caption = "Qunatity";
            oGrid.Columns.Item("U_IC").TitleObject.Caption = "ItemCode";
            oGrid.Columns.Item("U_IN").TitleObject.Caption = "Item Name";
            oGrid.Columns.Item("U_Model").TitleObject.Caption = "Model";
            oGrid.Columns.Item("U_Commodity").TitleObject.Caption = "Commodity";
            oGrid.Columns.Item("U_ContNo").TitleObject.Caption = "Container No.";
            oGrid.Columns.Item("U_ContType").TitleObject.Caption = "Container Type";
            oGrid.Columns.Item("U_IncTerms").TitleObject.Caption = "Incoterms Shipping Terms.";
            oGrid.Columns.Item("U_FeedPort").TitleObject.Caption = "Feeder Port";
            oGrid.Columns.Item("U_FeedETD").TitleObject.Caption = "Feeder ETD";
            oGrid.Columns.Item("U_MainPort").TitleObject.Caption = "Main Port";
            oGrid.Columns.Item("U_MainETD").TitleObject.Caption = "Main ETD";
            oGrid.Columns.Item("U_Pod").TitleObject.Caption = "Pod";
            oGrid.Columns.Item("U_PodETA").TitleObject.Caption = "POD ETA";
            oGrid.Columns.Item("U_GateDate").TitleObject.Caption = "Gate Out Date";
            oGrid.Columns.Item("U_EmptyDate").TitleObject.Caption = "Empty in Date";
            oGrid.Columns.Item("U_CNEE").TitleObject.Caption = "CNEE";
            oGrid.Columns.Item("U_Notify").TitleObject.Caption = "Notify";
            oGrid.Columns.Item("U_Forwarder").TitleObject.Caption = "Forwarder";
            oGrid.Columns.Item("U_ShipLine").TitleObject.Caption = "Shipping Line";
            oGrid.Columns.Item("U_FreeDays").TitleObject.Caption = "Free Days";
            oGrid.Columns.Item("U_Cost").TitleObject.Caption = "O/F Cost";
            oGrid.Columns.Item("U_Telex").TitleObject.Caption = "Telex";
            oGrid.Columns.Item("U_TelexDate").TitleObject.Caption = "Telex Date";
            oGrid.Columns.Item("Price").TitleObject.Caption = "UnitPrice";

            oGrid.Columns.Item("DocNum").Editable = false;
            oGrid.Columns.Item("Select").Editable = true;

            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            oGrid.AutoResizeColumns();
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None;
            AssignMatrixLineno(oGrid, aform);
            aform.PaneLevel = 3;
        }



        //    private void DataBind_ITems(SAPbouiCOM.Form aform)
        //    {
        //        try
        //        {
        //            aform.Freeze(true);
        //            string strDocNum = "0";
        //            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("1").Specific;

        //            for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
        //            {
        //                SAPbouiCOM.CheckBoxColumn oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");
        //                if (oCheckBox1.IsChecked(intRow))
        //                {
        //                    if (string.IsNullOrEmpty(strDocNum))
        //                        strDocNum = oGrid.DataTable.GetValue("DocNum", intRow).ToString();
        //                    else
        //                        strDocNum += "," + oGrid.DataTable.GetValue("DocNum", intRow).ToString();
        //                }
        //            }

        //            string aCardCode = GetCustomer(aform);
        //            string strQuery;
        //            oGrid = (SAPbouiCOM.Grid)aform.Items.Item("14").Specific;


        //            strQuery = string.Format(@"
        //        SELECT 'N' AS ""Select"", 
        //       T1.""DocEntry"", 
        //       T1.""DocNum"", 
        //       T1.""U_VC"", 
        //       T1.""U_VN"", 
        //       T0.""LineId"", 
        //       T0.""U_IC"", 
        //       T0.""U_IN"", 
        //       T0.""U_POQ"",
        //       T0.""U_Commodity"", T0.""U_Model"", T0.""U_ContNo"", T0.""U_ContType"", T0.""U_IncTerms"", T0.""U_FeedPort"", T0.""U_FeedETD"", T0.""U_MainPort"", T0.""U_MainETD"", T0.""U_Pod"", T0.""U_PodETA"", T0.""U_GateDate"", T0.""U_EmptyDate"", T0.""U_CNEE"", T0.""U_Notify"", T0.""U_Forwarder"", T0.""U_ShipLine"", T0.""U_FreeDays"", T0.""U_Cost"", T0.""U_Telex"", T0.""U_TelexDate""
        //FROM ""@BLC_BOLD"" T1 
        //INNER JOIN ""@BLC_BLD1"" T0 ON T1.""DocEntry"" = T0.""DocEntry"" 
        //WHERE T0.""DocEntry"" IN ({0}) and  (T0.""U_IC"" is not null  or  T0.""U_IC""='') ORDER BY T0.""DocEntry"", T0.""LineId""", strDocNum);

        //            oGrid.DataTable.ExecuteQuery(strQuery);


        //            oGrid.Columns.Item("U_VC").TitleObject.Caption = "Supplier Code";
        //            oEditTextColumn = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("U_VC");
        //            oEditTextColumn.LinkedObjectType = "2";

        //            oGrid.Columns.Item("U_VN").TitleObject.Caption = "Supplier Name";
        //            oGrid.Columns.Item("DocNum").TitleObject.Caption = "BL Doc. Number";
        //            oGrid.Columns.Item("LineId").TitleObject.Caption = "Line No.";
        //            oGrid.Columns.Item("U_POQ").TitleObject.Caption = "Qunatity";
        //            oGrid.Columns.Item("U_IC").TitleObject.Caption = "ItemCode";
        //            oGrid.Columns.Item("U_IN").TitleObject.Caption = "Item Name";
        //            oGrid.Columns.Item("U_Model").TitleObject.Caption = "Model";
        //            oGrid.Columns.Item("U_Commodity").TitleObject.Caption = "Commodity";
        //            oGrid.Columns.Item("U_ContNo").TitleObject.Caption = "Container No.";
        //            oGrid.Columns.Item("U_ContType").TitleObject.Caption = "Container Type";
        //            oGrid.Columns.Item("U_IncTerms").TitleObject.Caption = "Incoterms Shipping Terms.";
        //            oGrid.Columns.Item("U_FeedPort").TitleObject.Caption = "Feeder Port";
        //            oGrid.Columns.Item("U_FeedETD").TitleObject.Caption = "Feeder ETD";
        //            oGrid.Columns.Item("U_MainPort").TitleObject.Caption = "Main Port";
        //            oGrid.Columns.Item("U_MainETD").TitleObject.Caption = "Main ETD";
        //            oGrid.Columns.Item("U_Pod").TitleObject.Caption = "Pod";
        //            oGrid.Columns.Item("U_PodETA").TitleObject.Caption = "POD ETA";
        //            oGrid.Columns.Item("U_GateDate").TitleObject.Caption = "Gate Out Date";
        //            oGrid.Columns.Item("U_EmptyDate").TitleObject.Caption = "Empty in Date";
        //            oGrid.Columns.Item("U_CNEE").TitleObject.Caption = "CNEE";
        //            oGrid.Columns.Item("U_Notify").TitleObject.Caption = "Notify";
        //            oGrid.Columns.Item("U_Forwarder").TitleObject.Caption = "Forwarder";
        //            oGrid.Columns.Item("U_ShipLine").TitleObject.Caption = "Shipping Line";
        //            oGrid.Columns.Item("U_FreeDays").TitleObject.Caption = "Free Days";
        //            oGrid.Columns.Item("U_Cost").TitleObject.Caption = "O/F Cost";
        //            oGrid.Columns.Item("U_Telex").TitleObject.Caption = "Telex";
        //            oGrid.Columns.Item("U_TelexDate").TitleObject.Caption = "Telex Date";

        //            oGrid.Columns.Item("DocEntry").Visible = false;
        //            string[] readOnlyColumns = { "DocNum", "LineId", "U_VN", "U_VN", "U_IC", "U_IN","U_Commodity", "U_Model", "U_ContNo", "U_ContType", "U_IncTerms", "U_FeedPort", "U_FeedETD", "U_MainPort", "U_MainETD", "U_Pod", "U_PodETA", "U_GateDate", "U_EmptyDate", "U_CNEE", "U_Notify", "U_Forwarder", "U_ShipLine", "U_FreeDays", "U_Cost", "U_Telex", "U_TelexDate",
        //                         "U_POQ" };

        //            foreach (string column in readOnlyColumns)
        //            {
        //                oGrid.Columns.Item(column).Editable = false;
        //            }

        //            oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

        //            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

        //            oGrid.AutoResizeColumns();


        //            AssignMatrixLineno(oGrid, aform);
        //            aform.PaneLevel = 3;
        //            aform.Freeze(false);
        //        }
        //        catch (Exception ex)
        //        {
        //            EventHandler.oApplication.StatusBar.SetText("Item Copy from: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

        //            aform.Freeze(false);
        //        }
        //    }
        public void AssignMatrixLineno(SAPbouiCOM.Grid aGrid, SAPbouiCOM.Form aform)
        {
            aform.Freeze(true);
            try
            {
                for (int intNo = 0; intNo < aGrid.DataTable.Rows.Count; intNo++)
                {
                    aGrid.RowHeaders.SetText(intNo, (intNo + 1).ToString());
                }
            }
            catch (Exception ex)
            {
                // Handle exception if needed
            }
            aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#";
            aform.Freeze(false);
        }

        private string GetCustomer(SAPbouiCOM.Form aform)
        {
            string strBPCode = "'xxxxx'";
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("1000002").Specific;

            for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
            {
                if (!string.IsNullOrEmpty(oGrid.DataTable.GetValue("U_VC", intRow).ToString()))
                {
                    strBPCode += ",'" + oGrid.DataTable.GetValue("U_VC", intRow) + "'";
                }
            }

            return strBPCode;
        }
        //Get Document Numbers
        private string getDocNum(SAPbouiCOM.Form aform)
        {
            string strBPCode = "'xxxxx'";
            List<string> DocumentNo = new List<string>();
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("1000002").Specific;

            for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
            {
                SAPbouiCOM.CheckBoxColumn oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");

                if (!string.IsNullOrEmpty(oGrid.DataTable.GetValue("DocNum", intRow).ToString()) && oCheckBox1.IsChecked(intRow))
                {
                    DocumentNo.Add(oGrid.DataTable.GetValue("DocNum", intRow).ToString());
                    // strBPCode += ",'" + oGrid.DataTable.GetValue("DocNum", intRow) + "'";
                }
            }
            var distDocnum = DocumentNo.Distinct().ToList();
            strBPCode = string.Join(",", DocumentNo.Select(doc => $"'{doc}'"));
            //strBPCode = string.Join(",", distDocnum);
            return strBPCode;
        }
        private string GetItemsLists(SAPbouiCOM.Form aform)
        {
            string strBPCode = "'xxxxx'";
            List<string> DocumentNo = new List<string>();
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("1").Specific;

            for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
            {
                SAPbouiCOM.CheckBoxColumn oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");

                if (!string.IsNullOrEmpty(oGrid.DataTable.GetValue("DocNum", intRow).ToString()) && oCheckBox1.IsChecked(intRow))
                {
                    DocumentNo.Add(oGrid.DataTable.GetValue("DocNum", intRow).ToString());
                   // strBPCode += ",'" + oGrid.DataTable.GetValue("DocNum", intRow) + "'";
                }
            }
            var distDocnum = DocumentNo.Distinct().ToList();
            strBPCode = string.Join(",", DocumentNo.Select(doc => $"'{doc}'"));
            //strBPCode = string.Join(",", distDocnum);
            return strBPCode;
        }
        private string GetIContainerLists(SAPbouiCOM.Form aform)
        {
            string strContainerNo = "'xxxxx'";
            List<string> DocumentNo = new List<string>();
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aform.Items.Item("1").Specific;

            for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
            {
                SAPbouiCOM.CheckBoxColumn oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");
                string cont = oGrid.DataTable.GetValue("U_ContNo", intRow).ToString();
                if (!string.IsNullOrEmpty(oGrid.DataTable.GetValue("U_ContNo", intRow).ToString()) && oCheckBox1.IsChecked(intRow))
                {
                    DocumentNo.Add(oGrid.DataTable.GetValue("U_ContNo", intRow).ToString());
                   // strContainerNo += ",'" + oGrid.DataTable.GetValue("U_ContNo", intRow) + "'";
                }
            }
            strContainerNo = string.Join(",", DocumentNo.Select(doc => $"'{doc}'"));
           // strContainerNo = string.Join(",", DocumentNo);
            return strContainerNo;
        }

        public bool PopulatetoDocument(SAPbouiCOM.Form aForm)
        {
            try
            {
                string strDocNum = string.Empty;
                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)GlobalVariables.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)aForm.Items.Item("14").Specific;

                SAPbouiCOM.ProgressBar oPg = EventHandler.oApplication.StatusBar.CreateProgressBar("Data Copy In progress.....", oGrid.DataTable.Rows.Count, true);
                oPg.Text = "Data Copy In progress.....";
                List<CopyData> copyData = new List<CopyData>(); 
                for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
                {
                   // CopyData objCopy = new CopyData();
                    oPg.Value = oPg.Value + 1;
                    SAPbouiCOM.CheckBoxColumn oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");
                    if (oCheckBox1.IsChecked(intRow))
                    {
                        var existing = copyData.FirstOrDefault(x =>
                    x.ItemCode == oGrid.DataTable.GetValue("U_IC", intRow).ToString() 
    );
                       // if (existing == null)
                       // {
                            // If none found, create a new one
                            copyData.Add(new CopyData
                            {
                                ItemCode = oGrid.DataTable.GetValue("U_IC", intRow).ToString(),
                                Qty = Convert.ToDouble(oGrid.DataTable.GetValue("U_POQ", intRow).ToString()),
                                UnitPrice = Math.Round(Convert.ToDouble(oGrid.DataTable.GetValue("Price", intRow).ToString()), 2),
                                Warehouse = oGrid.DataTable.GetValue("U_Pod", intRow).ToString(),
                                BLNo = oGrid.DataTable.GetValue("DocNum", intRow).ToString(),
                                containerNo = oGrid.DataTable.GetValue("U_ContNo", intRow).ToString()
                                //containerNo = new List<string> { oGrid.DataTable.GetValue("U_ContNo", intRow).ToString() }
                            });
                        //}
                        //else
                        //{
                            
                        //    existing.Qty += Convert.ToDouble(oGrid.DataTable.GetValue("U_POQ", intRow));
                        //    if (!existing.containerNo.Contains(oGrid.DataTable.GetValue("U_ContNo", intRow).ToString()))
                        //    {
                        //        existing.containerNo.Add(oGrid.DataTable.GetValue("U_ContNo", intRow).ToString());
                        //    }
                        //}
                    }
                }
                oPg.Stop();

                oPg = EventHandler.oApplication.StatusBar.CreateProgressBar("Data Copy In progress.....", oGrid.DataTable.Rows.Count, true);
                oPg.Text = "Data Copy In progress.....";
                oGrid = (SAPbouiCOM.Grid)aForm.Items.Item("14").Specific;
                //for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
                //{
                //    oPg.Value = oPg.Value + 1;
                //    oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");
                //    if (oCheckBox1.IsChecked(intRow))
                //    {
                //        oRec.DoQuery("Update \"@BLC_BLD1\" set \"U_SelectFlag\"='Y' where \"DocEntry\"=" + oGrid.DataTable.GetValue("DocNum", intRow).ToString() + " and \"LineId\"=" + oGrid.DataTable.GetValue("LineId", intRow).ToString());
                //    }
                //    else
                //    {
                //        oRec.DoQuery("Update \"@BLC_BLD1\" set \"U_SelectFlag\"='N' where \"DocEntry\"=" + oGrid.DataTable.GetValue("DocNum", intRow).ToString() + " and \"LineId\"=" + oGrid.DataTable.GetValue("LineId", intRow).ToString());
                //    }
                //}
                oPg.Stop();

                if (copyData.Count > 0)
                {
                    GRPO oObj = new GRPO();
                    EventHandler.oApplication.StatusBar.SetText(
                                          "Copying data to Grpo. Please wait...",
                                          SAPbouiCOM.BoMessageTime.bmt_Short,
                                          SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    if (oObj.PopulateMultiplePODetails(copyData) == true)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.StatusBar.SetText("Populate Matrix Function: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                return false;
            }

        }


        #endregion
        public void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        oForm = EventHandler.oApplication.Forms.Item(GlobalVariables.BillWizardID);
                        if (pVal.BeforeAction == false)
                        {
                            if (pVal.ItemUID == "1000003")
                            {
                                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("1000002").Specific;
                                if (oGrid.DataTable.Rows.Count - 1 < 0)
                                {
                                    oGrid.DataTable.Rows.Add();
                                }
                                if (oGrid.DataTable.GetValue("CardCode", oGrid.DataTable.Rows.Count - 1).ToString() != "")
                                {
                                    oGrid.DataTable.Rows.Add();
                                }
                            }

                            if (pVal.ItemUID == "1000004")
                            {
                                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("1000002").Specific;
                                for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
                                {
                                    if (oGrid.Rows.IsSelected(intRow))
                                    {
                                        oGrid.DataTable.Rows.Remove(intRow);
                                        return; // Exit after removing row
                                    }
                                }
                            }

                            if (pVal.ItemUID == "15")
                            {
                                oForm.PaneLevel = oForm.PaneLevel - 1;
                            }



                            if (pVal.ItemUID == "16")
                            {
                                if (oForm.PaneLevel == 1)
                                {
                                    DataBind(oForm);
                                }
                                else if (oForm.PaneLevel == 2)
                                {
                                    DataBind_ITems(oForm);
                                }
                            }

                            if (pVal.ItemUID == "27")
                            {
                                if (EventHandler.SBO_Application.MessageBox("Do you want to Export the selected document details into Excel?", 1, "Continue", "Cancel") == 2)
                                {
                                    return;
                                }
                            }

                            if (pVal.ItemUID == "3")
                            {
                                if (EventHandler.oApplication.MessageBox("Do you want to copy the selected document details into Bill of Lading  Document?", 1, "Continue", "Cancel") == 2)
                                {
                                    return; 
                                }
                                if (PopulatetoDocument(oForm))
                                {
                                    oForm.Close();
                                }
                            }

                            if (pVal.ItemUID == "4")
                            {
                                SelectAll(oForm, true);
                            }

                            if (pVal.ItemUID == "5")
                            {
                                SelectAll(oForm, false);
                            }
                        }

                        break;

                }
            }
            catch (Exception ex)
            {
                EventHandler.oApplication.MessageBox(ex.Message);
            }
            finally
            {
            }
        }
        private void SelectAll(SAPbouiCOM.Form aForm, bool aChoice)
        {
            aForm.Freeze(true);

            SAPbouiCOM.Grid oGrid;
            if (aForm.PaneLevel == 3)
            {
                oGrid = (SAPbouiCOM.Grid)aForm.Items.Item("14").Specific;

            }
            else
            {
                oGrid = (SAPbouiCOM.Grid)aForm.Items.Item("1").Specific;

            }

            for (int intRow = 0; intRow < oGrid.DataTable.Rows.Count; intRow++)
            {
                SAPbouiCOM.CheckBoxColumn oCheckBox1 = (SAPbouiCOM.CheckBoxColumn)oGrid.Columns.Item("Select");
                oCheckBox1.Check(intRow, aChoice);
            }

            aForm.Freeze(false);
        }






    }
}
