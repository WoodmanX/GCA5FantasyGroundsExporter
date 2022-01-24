//These aren't usually required Imports within Visual Studio, but have to be included
//here now because the plugin compiler doesn't make these associations automatically.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;

//Everything from here and below is normal code and Imports and such, just as it is
//when developing within Visual Studio for VB projects.
using GCA5Engine;
using System.Drawing;
using System.Reflection;
using GCA5.Interfaces;

namespace GCA5FantasyGroundsExporter
{
    public class GCA5FantasyGroundsExporter : GCA5.Interfaces.IExportSheet
    {
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;

        public string PluginName()
        {
            return "Fantasy Grunds PC export";
        }

        public string PluginDescription()
        {
            return "Export Character as PC to Fantasy Grunds";
        }

        public void CreateOptions(SheetOptionsManager mySheetOptions)
        {

        }

        public bool GenerateExport(Party Party, string TargetFilename, SheetOptionsManager Options)
        {
            throw new NotImplementedException();
        }

        

        

        public string PluginVersion()
        {
            throw new NotImplementedException();
        }

        public int PreferredFilterIndex()
        {
            throw new NotImplementedException();
        }

        public bool PreviewOptions(SheetOptionsManager Options)
        {
            throw new NotImplementedException();
        }

        public string SupportedFileTypeFilter()
        {
            return "";
        }

        public void UpgradeOptions(SheetOptionsManager Options)
        {
            throw new NotImplementedException();
        }
    }
}
