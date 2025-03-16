using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
namespace BillOfLading
{

    /// <summary>
    /// GGlobally whatever variable do you want declare here 
    /// We can use any class and module from here  
    /// </summary>
    /// <remarks></remarks>
    static class GlobalVariables
    {

        #region " ... Common For SAP ..."
        public static SAPbobsCOM.Company oCompany;
        public static GlobalFunctions oGFun = new GlobalFunctions();
        public static SAPbouiCOM.Form oForm;
        public static List<BLData> dt = new List<BLData>();


        #endregion

        #region " ... Common For Forms ..."

        public static string bldID = "BLC_BOLD";
        public static string bldXML = "Presentation_Layer.Masters.Bill_of_lading.xml";
        public static billoflanding oBilloflanding = new billoflanding();



        public static GRPO ARinvoice = new GRPO();

        public static string BillWizardID = "frm_OBOL";
        public static string BillWizardXML = "Presentation_Layer.Masters.xml_Bill.xml";
        public static BillofLadingWizard oBillWizard = new BillofLadingWizard();

        public static SAPbouiCOM.Form frmAP;

        #endregion

        #region " ... Gentral Purpose ..."
        public static long v_RetVal;
        public static int v_ErrCode;
        public static string v_ErrMsg = "";
        public static string addonName = "BillOfLading";
        public static string sQuery = "";
        public static string BankFileName = "";
        public static string FileName = "";
        #endregion
    }
}
