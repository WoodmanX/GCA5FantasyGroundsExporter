//These aren't usually required Imports within Visual Studio, but have to be included
//here now because the plugin compiler doesn't make these associations automatically.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Security;

//Everything from here and below is normal code and Imports and such, just as it is
//when developing within Visual Studio for VB projects.
using GCA5Engine;
using System.Drawing;
using System.Globalization;
using System.Reflection;
using GCA5.Interfaces;

namespace GCA5FantasyGroundsExporterNPC
{
    public class GCA5FantasyGroundsExporterNPC
    {
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;

        private const string PLUGINVERSION = "1.1.0.0";
        private SheetOptionsManager myOptions;

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="Party"></param>
        /// <param name="TargetFilename"></param>
        /// <param name="Options"></param>
        /// <returns></returns>
        public bool GenerateExport(Party Party, string TargetFilename, SheetOptionsManager Options)
        {
            myOptions = Options;

            FileWriter ExportWriter = new FileWriter();
            ExportWriter.FileOpen(TargetFilename);
            ExportNPC(Party.Current, ExportWriter);

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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportNPC(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            fileWriter.Paragraph("<root release=\"4 | CoreRPG:3\" version=\"3.2\">");
            fileWriter.Paragraph("<npc>");
            //Name
            fileWriter.Paragraph(EscapedItem("name", "string", myCharacter.Name));

            ExportAbilities(myCharacter, fileWriter);
            Exportattributes(myCharacter, fileWriter);
            ExportEncumberance(myCharacter, fileWriter);
            ExportCombat(myCharacter, fileWriter);
            ExportTraits(myCharacter, fileWriter);
            ExportInventory(myCharacter, fileWriter);
            ExportPointTotals(myCharacter, fileWriter);
            fileWriter.Paragraph("</npc>");
            fileWriter.Paragraph("</root>");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportAbilities(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<abilities>");
            ExportSkills(myCharacter, fileWriter);
            fileWriter.Paragraph("</abilities>");

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportSkills(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var mySkills = myCharacter.ItemsByType[(int)TraitTypes.Skills];
            var mySpells = myCharacter.ItemsByType[(int)TraitTypes.Spells];
            int i = 1;

            fileWriter.Paragraph("<abilitieslist>");
            foreach (GCATrait skill in mySkills)
            {
                    var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                    fileWriter.Paragraph(index);
                    fileWriter.Paragraph(EscapedItem("name", "string", skill.FullNameTL));
                    fileWriter.Paragraph(EscapedItem("level", "number", skill.Level.ToString()));
                    fileWriter.Paragraph(index.Insert(1, "/"));
                    i++;
            }

            foreach (GCATrait spell in mySpells)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                var myClass = spell.get_TagItem("class");
                var myResist = "";
                var myCastingTime = spell.get_TagItem("time");


                if (myClass.Contains("/"))
                {
                    var mySplit = myClass.Split('/');
                    myClass = mySplit[0];
                    myResist = mySplit[1];
                }

                fileWriter.Paragraph(index);
                fileWriter.Paragraph(EscapedItem("name", "string", "Spell:" + spell.FullNameTL));
                fileWriter.Paragraph(EscapedItem("level", "number", spell.Level.ToString()));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            fileWriter.Paragraph("</abilitieslist>");
            }

        private string EscapedItem(string tagName, string tagType, string item)
        {
            return "<" + tagName + " type=\"" + tagType + "\">" + SecurityElement.Escape(item) + "</" + tagName + ">";
        }

        /// <summary>
        /// utility function to check if a given item is hidden
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool IsItemHidden(GCATrait item)
        {
            return !(item.get_TagItem("hidden") == "");
        }
        /// <summary>
        /// assembles the damage string
        /// </summary>
        /// <param name="item"></param>
        /// <param name="mode"></param>
        /// <returns></returns>
        private string GetDamageString(GCATrait item, int mode)
        {
            return item.DamageModeTagItem(mode, "chardamage") + " " + item.DamageModeTagItem(mode, "chardamtype");
        }
    }
