﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace CREPE
{
    public partial class rubanAddConge
    {
        private void rubanAddConge_Load(object sender, RibbonUIEventArgs e)
        {

        }



        private void AjoutConge_Click(object sender, RibbonControlEventArgs e)
        {

            Globals.ThisAddIn.creationDeConge();
        }
    }
}
