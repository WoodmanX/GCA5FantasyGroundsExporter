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
    public class GCA5FantasyGroundsExporterNPC : GCA5.Interfaces.IExportSheet
    {
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;

        private const string PLUGINVERSION = "1.1.0.0";
        private SheetOptionsManager myOptions;

        public string PluginName()
        {
            return "Fantasy Grounds NPC export";
        }

        public string PluginDescription()
        {
            return "Export Character as NPC to Fantasy Grounds for more information see https://github.com/WoodmanX/GCA5FantasyGroundsExporter";
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

            ExportNotes(myCharacter, fileWriter);
            ExportAbilities(myCharacter, fileWriter);
            ExportTraits(myCharacter, fileWriter);

            Exportattributes(myCharacter, fileWriter);
            ExportCombat(myCharacter, fileWriter);
 
            fileWriter.Paragraph("</npc>");
            fileWriter.Paragraph("</root>");
        }

        private void ExportNotes(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<notes type=\"formattedtext\">");
            fileWriter.Paragraph(EscapedItem("p","Race: " + myCharacter.Race));
            fileWriter.Paragraph(EscapedItem("p", "height: " + myCharacter.Height));
            fileWriter.Paragraph(EscapedItem("p", "weight: " + myCharacter.Weight));
            fileWriter.Paragraph(EscapedItem("p", "age: " + myCharacter.Age));
            fileWriter.Paragraph(EscapedItem("p", "appearance: " + myCharacter.Appearance));
            fileWriter.Paragraph("</notes>");
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

        private void ExportTraits(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var x = myCharacter.ItemsByName("Size Modifier", (int)TraitTypes.Attributes);
            var sm = "0";
            if (x.Count > 0)
            {
                GCATrait gCATrait = (GCATrait)x[1];
                sm = gCATrait.Score.ToString();
            }

            fileWriter.Paragraph("<traits>");
            fileWriter.Paragraph(EscapedItem("sizemodifier", "string", sm));
            ExportReactionMods(myCharacter, fileWriter);
            ExportAdsnDisads(myCharacter, fileWriter);
            fileWriter.Paragraph("</traits>");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportReactionMods(GCACharacter myCharacter, FileWriter fileWriter)
        {
            GCATrait reaction = myCharacter.ItemByNameAndExt("Reaction", (int)TraitTypes.Attributes);
            string reactionmods = reaction.get_TagItem("bonuslist");
            reactionmods = reactionmods + ", " + reaction.get_TagItem("conditionallist");

            fileWriter.Paragraph(EscapedItem("reactionmodifiers", "string", reactionmods));
        }

        private void ExportAdsnDisads(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var Templates = myCharacter.ItemsByType[(int)TraitTypes.Templates];
            var Ads = myCharacter.ItemsByType[(int)TraitTypes.Advantages];
            var Perks = myCharacter.ItemsByType[(int)TraitTypes.Perks];
            var Features = myCharacter.ItemsByType[(int)TraitTypes.Features];
            var Disads = myCharacter.ItemsByType[(int)TraitTypes.Disadvantages];
            var Quirks = myCharacter.ItemsByType[(int)TraitTypes.Quirks];

            string AdsnDisads = "";
            foreach (GCATrait item in Templates)
            {
                AdsnDisads = AdsnDisads + item.FullName + "\\r";
            }

            foreach (GCATrait item in Ads)
            {
                AdsnDisads = AdsnDisads + item.FullName + "\\r";
            }

            foreach (GCATrait item in Perks)
            {
                AdsnDisads = AdsnDisads + item.FullName + "\\r";
            }

            foreach (GCATrait item in Features)
            {
                AdsnDisads = AdsnDisads + item.FullName + "\\r";
            }

            foreach (GCATrait item in Disads)
            {
                AdsnDisads = AdsnDisads + item.FullName + "\\r";
            }

            foreach (GCATrait item in Quirks)
            {
                AdsnDisads = AdsnDisads + item.FullName + "\\r";
            }

            fileWriter.Paragraph("<description type=\"string\">");
            fileWriter.Paragraph(SecurityElement.Escape(AdsnDisads));
            fileWriter.Paragraph("</description>");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void Exportattributes(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<attributes>");

            fileWriter.Paragraph(EscapedItem("strength", "number", myCharacter.ItemByNameAndExt("ST", (int)TraitTypes.Stats).Score.ToString()));            
            fileWriter.Paragraph(EscapedItem("dexterity", "number", myCharacter.ItemByNameAndExt("DX", (int)TraitTypes.Stats).Score.ToString()));            
            fileWriter.Paragraph(EscapedItem("intelligence", "number", myCharacter.ItemByNameAndExt("IQ", (int)TraitTypes.Stats).Score.ToString()));            
            fileWriter.Paragraph(EscapedItem("health", "number", myCharacter.ItemByNameAndExt("HT", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(EscapedItem("hitpoints", "number", myCharacter.ItemByNameAndExt("Hit Points", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(EscapedItem("will", "number", myCharacter.ItemByNameAndExt("Will", (int)TraitTypes.Stats).Score.ToString()));            
            fileWriter.Paragraph(EscapedItem("perception", "number", myCharacter.ItemByNameAndExt("Perception", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(EscapedItem("fatiguepoints", "number", myCharacter.ItemByNameAndExt("Fatigue Points", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(EscapedItem("basicspeed", "string", myCharacter.ItemByNameAndExt("Basic Speed", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(EscapedItem("thrust", "string", myCharacter.BaseTH));
            fileWriter.Paragraph(EscapedItem("swing", "string", myCharacter.BaseSW));                       
            fileWriter.Paragraph("</attributes>");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportCombat(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<combat>");
            fileWriter.Paragraph(EscapedItem("dr", "string", getDR(myCharacter)));
            fileWriter.Paragraph(EscapedItem("dodge", "number", myCharacter.ItemByNameAndExt("Dodge", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(EscapedItem("parry", "number", myCharacter.ParryScore.ToString()));
            fileWriter.Paragraph(EscapedItem("block", "number", myCharacter.BlockScore.ToString()));
            ExportMeleeList(myCharacter, fileWriter);
            ExportRangedList(myCharacter, fileWriter);

            fileWriter.Paragraph("</combat>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportMeleeList(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var attackIndex = 1;

            fileWriter.Paragraph("<meleecombatlist>");

            foreach (GCATrait Item in myCharacter.Items)
            {
                if (Item.DamageModeTagItemCount("charreach") > 0)
                {
                    var ModeCount = Item.DamageModeTagItemCount("charreach");

                    if (!IsItemHidden(Item) && ModeCount > 0)
                    {
                        var curMode = 1;
                        var index = "<id-" + attackIndex.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                        fileWriter.Paragraph(index);

                        fileWriter.Paragraph(EscapedItem("name", "string", Item.Name));
                        fileWriter.Paragraph(EscapedItem("st", "string", Item.DamageModeTagItem(curMode, "charminst")));
                        fileWriter.Paragraph(EscapedItem("cost", "string", Item.get_TagItem("cost")));
                        fileWriter.Paragraph(EscapedItem("weight", "string", Item.get_TagItem("weight")));
                        fileWriter.Paragraph(EscapedItem("text", "string", Item.get_TagItem("description")));
                        fileWriter.Paragraph(EscapedItem("tl", "string", Item.get_TagItem("techlvl")));

                        fileWriter.Paragraph("<meleemodelist>");

                        do
                        {

                            var indexMode = "<id-" + (curMode).ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                            fileWriter.Paragraph(indexMode);
                            fileWriter.Paragraph(EscapedItem("name", "string", Item.DamageModeTagItem(curMode, "name")));
                            fileWriter.Paragraph(EscapedItem("level", "number", Item.DamageModeTagItem(curMode, "charskillscore")));
                            fileWriter.Paragraph(EscapedItem("damage", "string", GetDamageString(Item, curMode)));
                            fileWriter.Paragraph(EscapedItem("reach", "string", Item.DamageModeTagItem(curMode, "charreach")));
                            fileWriter.Paragraph(EscapedItem("parry", "string", Item.DamageModeTagItem(curMode, "parry")));
                            fileWriter.Paragraph(indexMode.Insert(1, "/"));

                            curMode = Item.DamageModeTagItemAt("charreach", curMode + 1);
                        } while (curMode > 0);

                        fileWriter.Paragraph("</meleemodelist>");
                        fileWriter.Paragraph(index.Insert(1, "/"));
                        attackIndex++;
                    }


                }
            }

            fileWriter.Paragraph("</meleecombatlist>");
        }

        /// <summary>
        /// exports every ranged attackmode of the given cahracter
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportRangedList(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var attackIndex = 1;

            fileWriter.Paragraph("<rangedcombatlist>");

            //As a lot of things can give you a ranged attack we have to iterate over all items
            foreach (GCATrait Item in myCharacter.Items)
            {
                //Every ranged attack has at least one entry with a range
                var ModeCount = Item.DamageModeTagItemCount("charrangemax");

                if (!IsItemHidden(Item) && ModeCount > 0)
                {
                    var curMode = 1;
                    var index = "<id-" + attackIndex.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                    fileWriter.Paragraph(index);
                    fileWriter.Paragraph(EscapedItem("name", "string", Item.Name));
                    fileWriter.Paragraph(EscapedItem("st", "string", Item.DamageModeTagItem(curMode, "charminst")));
                    fileWriter.Paragraph(EscapedItem("bulk", "number", Item.DamageModeTagItem(curMode, "bulk")));
                    fileWriter.Paragraph(EscapedItem("lc", "string", Item.DamageModeTagItem(curMode, "lc")));
                    fileWriter.Paragraph(EscapedItem("text", "string", Item.get_TagItem("description")));
                    fileWriter.Paragraph(EscapedItem("tl", "string", Item.get_TagItem("techlvl")));

                    fileWriter.Paragraph("<rangedmodelist>");
                    do
                    {
                        var indexMode = "<id-" + (curMode).ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                        fileWriter.Paragraph(indexMode);
                        fileWriter.Paragraph(EscapedItem("name", "string", Item.DamageModeTagItem(curMode, "name")));
                        fileWriter.Paragraph(EscapedItem("level", "number", Item.DamageModeTagItem(curMode, "charskillscore")));
                        fileWriter.Paragraph(EscapedItem("damage", "string", GetDamageString(Item, curMode)));
                        fileWriter.Paragraph(EscapedItem("acc", "number", Item.DamageModeTagItem(curMode, "acc")));
                        var range = Item.DamageModeTagItem(curMode, "charrangehalfdam") + "/" + Item.DamageModeTagItem(curMode, "charrangemax");
                        fileWriter.Paragraph(EscapedItem("range", "string", range));
                        fileWriter.Paragraph(EscapedItem("rof", "string", Item.DamageModeTagItem(curMode, "rof")));
                        fileWriter.Paragraph(EscapedItem("shots", "string", Item.DamageModeTagItem(curMode, "shots")));
                        fileWriter.Paragraph(EscapedItem("rcl", "number", Item.DamageModeTagItem(curMode, "rcl")));
                        fileWriter.Paragraph(indexMode.Insert(1, "/"));

                        curMode = Item.DamageModeTagItemAt("charrangemax", curMode + 1);
                    } while (curMode > 0);
                    fileWriter.Paragraph("</rangedmodelist>");
                    fileWriter.Paragraph(index.Insert(1, "/"));
                    attackIndex++;
                }
            }

            fileWriter.Paragraph("</rangedcombatlist>");
        }


        private string getDR(GCACharacter myCharacter)
        {
            var myDr = "0*";

            foreach (string myname in myCharacter.Body.AllVisibleBodyPartNames())
            {
                if (myname == "Torso")
                {
                    BodyItem myBodyItem = myCharacter.Body.Item(myname);
                    myDr = myBodyItem.DR;
                }
            }

            return myDr;
           
        }
        private string EscapedItem(string tagName, string tagType, string item)
        {
            return "<" + tagName + " type=\"" + tagType + "\">" + SecurityElement.Escape(item) + "</" + tagName + ">";
        }

        private string EscapedItem(string tagName, string item)
        {
            return "<" + tagName + ">" + SecurityElement.Escape(item) + "</" + tagName + ">";
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
}
