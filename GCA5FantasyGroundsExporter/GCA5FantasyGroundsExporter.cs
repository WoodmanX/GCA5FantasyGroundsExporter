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
using System.Globalization;
using System.Reflection;
using GCA5.Interfaces;

namespace GCA5FantasyGroundsExporter
{
    public class GCA5FantasyGroundsExporter : GCA5.Interfaces.IExportSheet
    {
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;

        private const string PLUGINVERSION = "1.0.0.1";
        private SheetOptionsManager myOptions;
        //private List<Skill> Skills;

        public string PluginName()
        {
            return "Fantasy Grunds PC export";
        }

        public string PluginDescription()
        {
            return "Export Character as PC to Fantasy Grunds";
        }

        public string PluginVersion()
        {
            return PLUGINVERSION;
        }

        public void UpgradeOptions(SheetOptionsManager Options)
        {
            //Not needed as of now
        }

        public string SupportedFileTypeFilter()
        {
            return "XML files (*.xml)|*.xml";
        }
        public int PreferredFilterIndex()
        {
            //Only xml files are supported
            return 0;
        }

        public void CreateOptions(SheetOptionsManager mySheetOptions)
        {
            SheetOption mySheetOption = new SheetOption();
            SheetOptionDisplayFormat myDisplayFormat = new SheetOptionDisplayFormat();

            myDisplayFormat.BackColor = SystemColors.Info;
            myDisplayFormat.CaptionLocalBackColor = SystemColors.Info;

            mySheetOption.Clear();
            mySheetOption.Name = "Header_Description";
            mySheetOption.Type = GCA5Engine.OptionType.Header;
            mySheetOption.UserPrompt = PluginName() + PluginVersion();
            mySheetOption.DisplayFormat = myDisplayFormat;
            mySheetOptions.AddOption(mySheetOption);

            mySheetOption.Clear();
            mySheetOption.Name = "Description";
            mySheetOption.Type = GCA5Engine.OptionType.Caption;
            mySheetOption.UserPrompt = PluginDescription();
            mySheetOption.DisplayFormat = myDisplayFormat;
            mySheetOptions.AddOption(mySheetOption);

        }

        public bool PreviewOptions(SheetOptionsManager Options)
        {
            // As of now its not needed
            return true;
        }


        public bool GenerateExport(Party Party, string TargetFilename, SheetOptionsManager Options)
        {
            myOptions = Options;

            FileWriter ExportWriter = new FileWriter( );
            ExportWriter.FileOpen(TargetFilename);
            ExportPC(Party.Current, ExportWriter);

            if (Party.Characters.Count > 20)
            {
                DialogOptions_RequestedOptions e = new DialogOptions_RequestedOptions();
                SheetOptionsManager SOM = new GCA5Engine.SheetOptionsManager("RunSpecificOptions For " + PluginName());
                e.RunSpecificOptions = SOM;
                RequestRunSpecificOptions.Invoke(this, e);
            }

            try
            {
                ExportWriter.FileClose();
            }
            catch
            {

            }

            return true;
        }

        private void ExportPC(GCACharacter myCharacter, FileWriter fileWriter )
        {
            fileWriter.Paragraph("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            fileWriter.Paragraph("<root release=\"4 | CoreRPG:3\" version=\"3.2\">");
            fileWriter.Paragraph("<character>");
            //Name
            fileWriter.Paragraph("<name type=\"string\">" + myCharacter.Name + "</name>");

            exportAbilities(myCharacter, fileWriter);

            fileWriter.Paragraph("</character>");
            fileWriter.Paragraph("</root>");
        }

        private void exportAbilities(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<abilities>");
            exportSkills(myCharacter, fileWriter);
            exportSpells(myCharacter, fileWriter);
            fileWriter.Paragraph("</abilities>");
        }

        private void exportSkills(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var mySkills = myCharacter.ItemsByType[(int)TraitTypes.Skills];
            int i = 1;

            fileWriter.Paragraph("<skilllist>");
            foreach (GCATrait skill in mySkills)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(index);
                fileWriter.Paragraph("<name type = \"string\">" + skill.FullNameTL + "</name>");
                fileWriter.Paragraph("<type type=\"string\""+ skill.SkillType + "</type>");
                fileWriter.Paragraph("<level type=\"number\"" + skill.Level + "</level>");
                fileWriter.Paragraph("<relativelevel type=\"string\"" + skill.RelativeLevel + "</relativelevel>");
                fileWriter.Paragraph("<points type=\"number\"" + skill.Points + "</points>");
                fileWriter.Paragraph("<text type=\"string\"" + skill.Notes + "</text>");
                fileWriter.Paragraph(index.Insert(1,"/"));
                i++;
            }

            fileWriter.Paragraph("</skilllist>");
        }

        private void exportSpells(GCACharacter myCharacter, FileWriter fileWriter)
        {

        }

    }

}
