using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace DataExtractor3
{
    public class Button1 : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public Button1()
        {
        }

        protected override void OnClick()
        {
            frmDataExtractor frmMyForm;
            frmMyForm = new frmDataExtractor();
            frmMyForm.Show();
            ArcMap.Application.CurrentTool = null;
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }
    }

}
