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

namespace GCA5FantasyGroundsExporter
{
    /// <summary>
    /// 
    /// </summary>
    public class GCA5FantasyGroundsExporter : GCA5.Interfaces.IExportSheet
    {
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;

        private const string PLUGINVERSION = "1.0.0.5";
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportPC(GCACharacter myCharacter, FileWriter fileWriter )
        {
            fileWriter.Paragraph("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            fileWriter.Paragraph("<root release=\"4 | CoreRPG:3\" version=\"3.2\">");
            fileWriter.Paragraph("<character>");
            //Name
            fileWriter.Paragraph( escapedItem("name", "string", myCharacter.Name));

            ExportAbilities(myCharacter, fileWriter);
            Exportattributes(myCharacter, fileWriter);
            ExportEncumberance(myCharacter, fileWriter);
            ExportCombat(myCharacter, fileWriter);
            exportTraits(myCharacter, fileWriter);
            fileWriter.Paragraph("</character>");
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
            ExportSpells(myCharacter, fileWriter);
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
            int i = 1;

            fileWriter.Paragraph("<skilllist>");
            foreach (GCATrait skill in mySkills)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(index);
                fileWriter.Paragraph(escapedItem("name", "string", skill.FullNameTL));
                fileWriter.Paragraph(escapedItem("type", "string", skill.SkillType));
                fileWriter.Paragraph(escapedItem("level", "number", skill.Level.ToString()));
                fileWriter.Paragraph(escapedItem("relativelevel", "string", skill.RelativeLevel));
                fileWriter.Paragraph(escapedItem("points", "number", skill.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", skill.Notes));
                fileWriter.Paragraph(index.Insert(1,"/"));
                i++;
            }

            fileWriter.Paragraph("</skilllist>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportSpells(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var mySpells = myCharacter.ItemsByType[(int)TraitTypes.Spells];
            int i = 1;

            fileWriter.Paragraph("<spelllist>");
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
                fileWriter.Paragraph(escapedItem("name", "string", spell.FullNameTL));
                fileWriter.Paragraph(escapedItem("level", "number", spell.Level.ToString()));
                fileWriter.Paragraph(escapedItem("class", "string", myClass));
                fileWriter.Paragraph(escapedItem("type", "string", spell.get_TagItem("type")));
                fileWriter.Paragraph(escapedItem("points", "number", spell.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", spell.Notes));
                fileWriter.Paragraph(escapedItem("time", "string", spell.get_TagItem("time")));
                fileWriter.Paragraph(escapedItem("duration", "string", spell.get_TagItem("duration")));
                fileWriter.Paragraph(escapedItem("costmaintain", "string", spell.get_TagItem("castingcost")));
                fileWriter.Paragraph(escapedItem("resist", "string", myResist));
                fileWriter.Paragraph(escapedItem("college", "string", spell.get_TagItem("cat")));
                fileWriter.Paragraph(escapedItem("page", "string", spell.get_TagItem("page")));                  

                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            fileWriter.Paragraph("</spelllist>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void Exportattributes(GCACharacter myCharacter, FileWriter fileWriter)
        {
            fileWriter.Paragraph("<attributes>");

            fileWriter.Paragraph(escapedItem("strength", "number", myCharacter.ItemByNameAndExt("ST", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("strength_points", "number", myCharacter.ItemByNameAndExt("ST", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("dexterity", "number", myCharacter.ItemByNameAndExt("DX", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("dexterity_points", "number", myCharacter.ItemByNameAndExt("DX", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("intelligence", "number", myCharacter.ItemByNameAndExt("IQ", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("intelligence_points", "number", myCharacter.ItemByNameAndExt("IQ", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("health", "number", myCharacter.ItemByNameAndExt("HT", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("health_points", "number", myCharacter.ItemByNameAndExt("HT", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("hitpoints", "number", myCharacter.ItemByNameAndExt("Hit Points", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("hitpoints_points", "number", myCharacter.ItemByNameAndExt("Hit Points", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("hps", "number", myCharacter.ItemByNameAndExt("Hit Points", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("will", "number", myCharacter.ItemByNameAndExt("Will", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("will_points", "number", myCharacter.ItemByNameAndExt("Will", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("perception", "number", myCharacter.ItemByNameAndExt("Perception", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("perception_points", "number", myCharacter.ItemByNameAndExt("Perception", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("fatiguepoints", "number", myCharacter.ItemByNameAndExt("Fatigue Points", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("fatiguepoints_points", "number", myCharacter.ItemByNameAndExt("Fatigue Points", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("fps", "number", myCharacter.ItemByNameAndExt("Fatigue Points", (int)TraitTypes.Stats).Score.ToString()));
            
            fileWriter.Paragraph(escapedItem("basiclift", "string", myCharacter.ItemByNameAndExt("Basic Lift", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("thrust", "string", myCharacter.BaseTH));
            fileWriter.Paragraph(escapedItem("swing", "string", myCharacter.BaseSW));
            fileWriter.Paragraph(escapedItem("basicspeed", "string", myCharacter.ItemByNameAndExt("Basic Speed", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("basicspeed_points", "number", myCharacter.ItemByNameAndExt("Basic Speed", (int)TraitTypes.Stats).Points.ToString()));

            var myBasicMove = myCharacter.ItemByNameAndExt("Basic Move", (int)TraitTypes.Stats).Score;
            fileWriter.Paragraph(escapedItem("basicmove", "string", myBasicMove.ToString()));
            fileWriter.Paragraph(escapedItem("basicmove_points", "number", myCharacter.ItemByNameAndExt("Basic Move", (int)TraitTypes.Stats).Points.ToString()));
            fileWriter.Paragraph(escapedItem("move", "string", myBasicMove.ToString()));

            fileWriter.Paragraph("</attributes>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportEncumberance(GCACharacter myCharacter, FileWriter fileWriter)
        {


            var curEnc = myCharacter.EncumbranceLevel;
            var noEnc = myCharacter.ItemByNameAndExt("No Encumbrance", (int)TraitTypes.Stats).DisplayScore;
            var ligEnc = myCharacter.ItemByNameAndExt("Light Encumbrance", (int)TraitTypes.Stats).DisplayScore;
            var medEnc = myCharacter.ItemByNameAndExt("Medium Encumbrance", (int)TraitTypes.Stats).DisplayScore;
            var heaEnc = myCharacter.ItemByNameAndExt("Heavy Encumbrance", (int)TraitTypes.Stats).DisplayScore;
            var xheEnc = myCharacter.ItemByNameAndExt("X-Heavy Encumbrance", (int)TraitTypes.Stats).DisplayScore;

            var noEncMov = myCharacter.ItemByNameAndExt("No Encumbrance Move", (int)TraitTypes.Stats).DisplayScore;
            var ligEncMov = myCharacter.ItemByNameAndExt("Light Encumbrance Move", (int)TraitTypes.Stats).DisplayScore;
            var medEncMov = myCharacter.ItemByNameAndExt("Medium Encumbrance Move", (int)TraitTypes.Stats).DisplayScore;
            var heaEncMov = myCharacter.ItemByNameAndExt("Heavy Encumbrance Move", (int)TraitTypes.Stats).DisplayScore;
            var xheEncMov = myCharacter.ItemByNameAndExt("X-Heavy Encumbrance Move", (int)TraitTypes.Stats).DisplayScore;

            var noEncDodge = (int)myCharacter.ItemByNameAndExt("Dodge", (int)TraitTypes.Stats).Score;

            int[] myEnc = {0,0,0,0,0};
            myEnc[curEnc] = 1;

            

            fileWriter.Paragraph("<encumbrance>");

            for(int i = 0; i < myEnc.Length; i++)
            {
                fileWriter.Paragraph(escapedItem("enc_"+i, "number", myEnc[i].ToString()));
            }

            fileWriter.Paragraph(escapedItem("enc0_weight", "string", noEnc.ToString()));
            fileWriter.Paragraph(escapedItem("enc1_weight", "string", ligEnc.ToString()));
            fileWriter.Paragraph(escapedItem("enc2_weight", "string", medEnc.ToString()));
            fileWriter.Paragraph(escapedItem("enc3_weight", "string", heaEnc.ToString()));
            fileWriter.Paragraph(escapedItem("enc4_weight", "string", xheEnc.ToString()));

            fileWriter.Paragraph(escapedItem("enc0_move", "string", noEncMov.ToString()));
            fileWriter.Paragraph(escapedItem("enc1_move", "string", ligEncMov.ToString()));
            fileWriter.Paragraph(escapedItem("enc2_move", "string", medEncMov.ToString()));
            fileWriter.Paragraph(escapedItem("enc3_move", "string", heaEncMov.ToString()));
            fileWriter.Paragraph(escapedItem("enc4_move", "string", xheEncMov.ToString()));

            fileWriter.Paragraph(escapedItem("enc0_dodge", "number", noEncDodge.ToString("D")));
            fileWriter.Paragraph(escapedItem("enc1_dodge", "number", (noEncDodge - 1).ToString("D")));
            fileWriter.Paragraph(escapedItem("enc2_dodge", "number", (noEncDodge - 2).ToString("D")));
            fileWriter.Paragraph(escapedItem("enc3_dodge", "number", (noEncDodge - 3).ToString("D")));
            fileWriter.Paragraph(escapedItem("enc4_dodge", "number", (noEncDodge - 4).ToString("D")));

            fileWriter.Paragraph("</encumbrance>");

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void ExportCombat(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var noEncDodge = (int)myCharacter.ItemByNameAndExt("Dodge", (int)TraitTypes.Stats).Score;

            fileWriter.Paragraph("<combat>");
            fileWriter.Paragraph(escapedItem("enc0_dodge", "number", noEncDodge.ToString("D")));
            fileWriter.Paragraph(escapedItem("enc1_dodge", "number", (noEncDodge - 1).ToString("D")));
            fileWriter.Paragraph(escapedItem("enc2_dodge", "number", (noEncDodge - 2).ToString("D")));
            fileWriter.Paragraph(escapedItem("enc3_dodge", "number", (noEncDodge - 3).ToString("D")));
            fileWriter.Paragraph(escapedItem("enc4_dodge", "number", (noEncDodge - 4).ToString("D")));

            fileWriter.Paragraph(escapedItem("dodge", "number", myCharacter.ItemByNameAndExt("Dodge", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("parry", "number", myCharacter.ItemByNameAndExt("Parry", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("block", "number", myCharacter.ItemByNameAndExt("Block", (int)TraitTypes.Stats).Score.ToString()));
            fileWriter.Paragraph(escapedItem("dr", "string", myCharacter.ItemByNameAndExt("DR", (int)TraitTypes.Stats).Score.ToString()));

            fileWriter.Paragraph("<protectionlist>");

            //iterate over all BodyItems of the current character
            int i = 1;
            foreach (string myBodyItemName in myCharacter.Body.AllVisibleBodyPartNames())
            {
                BodyItem myBodyItem = myCharacter.Body.Item(myBodyItemName);
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(index);
                fileWriter.Paragraph(escapedItem("location", "string", myBodyItem.Name));
                fileWriter.Paragraph(escapedItem("dr", "string", myBodyItem.DR));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }
            fileWriter.Paragraph("</protectionlist>");

            exportMeleeList(myCharacter, fileWriter);
            exportRangedList(myCharacter, fileWriter);

            fileWriter.Paragraph("</combat>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportMeleeList(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var attackIndex = 1;

            fileWriter.Paragraph("<meleecombatlist>");

            foreach (GCATrait Item in myCharacter.Items)
            {
                if (Item.DamageModeTagItemCount("charreach") > 0)
                {
                    var ModeCount = Item.DamageModeTagItemCount("charreach");

                    if (!isItemHidden(Item) && ModeCount > 0)
                    {
                        var curMode = 1;
                        var index = "<id-" + attackIndex.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                        
                        fileWriter.Paragraph(index);

                        fileWriter.Paragraph(escapedItem("name", "string", Item.Name));
                        fileWriter.Paragraph(escapedItem("st", "string", Item.DamageModeTagItem(curMode, "charminst")));
                        fileWriter.Paragraph(escapedItem("cost", "string", Item.get_TagItem("cost")));
                        fileWriter.Paragraph(escapedItem("weight", "string", Item.get_TagItem("weight")));
                        fileWriter.Paragraph(escapedItem("text", "string", Item.get_TagItem("description")));
                        fileWriter.Paragraph(escapedItem("tl", "string", Item.get_TagItem("techlvl")));

                        fileWriter.Paragraph("<meleemodelist>");

                        do
                        {
                            
                            var indexMode = "<id-" + (curMode).ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                            fileWriter.Paragraph(indexMode);
                            fileWriter.Paragraph(escapedItem("name", "string", Item.DamageModeTagItem(curMode, "name")));
                            fileWriter.Paragraph(escapedItem("level", "number", Item.DamageModeTagItem(curMode, "charskillscore")));
                            fileWriter.Paragraph(escapedItem("damage", "string", getDamageString(Item, curMode)));
                            fileWriter.Paragraph(escapedItem("reach", "string", Item.DamageModeTagItem(curMode, "charreach")));
                            fileWriter.Paragraph(escapedItem("parry", "string", Item.DamageModeTagItem(curMode, "parry")));
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
        private void exportRangedList(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var attackIndex = 1;

            fileWriter.Paragraph("<rangedcombatlist>");

            foreach (GCATrait Item in myCharacter.Items)
            {
                var ModeCount = Item.DamageModeTagItemCount("charrangemax");

                if (!isItemHidden(Item) && ModeCount > 0)
                {
                    var curMode = 1;
                    var index = "<id-" + attackIndex.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                    
                    fileWriter.Paragraph(index);
                    fileWriter.Paragraph(escapedItem("name", "string", Item.Name));
                    fileWriter.Paragraph(escapedItem("st", "string", Item.DamageModeTagItem(curMode, "charminst")));
                    fileWriter.Paragraph(escapedItem("bulk", "number", Item.DamageModeTagItem(curMode, "bulk")));
                    fileWriter.Paragraph(escapedItem("lc", "string", Item.DamageModeTagItem(curMode, "lc")));
                    fileWriter.Paragraph(escapedItem("text", "string", Item.get_TagItem("description")));
                    fileWriter.Paragraph(escapedItem("tl", "string", Item.get_TagItem("techlvl")));

                    fileWriter.Paragraph("<rangedmodelist>");
                    do
                    {
                        var indexMode = "<id-" + (curMode).ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";

                        fileWriter.Paragraph(indexMode);
                        fileWriter.Paragraph(escapedItem("name", "string", Item.DamageModeTagItem(curMode, "name")));
                        fileWriter.Paragraph(escapedItem("level", "number", Item.DamageModeTagItem(curMode, "charskillscore")));
                        fileWriter.Paragraph(escapedItem("damage", "string", getDamageString(Item, curMode)));
                        fileWriter.Paragraph(escapedItem("acc", "number", Item.DamageModeTagItem(curMode, "acc")));
                        var range = Item.DamageModeTagItem(curMode, "charrangehalfdam") + "/" + Item.DamageModeTagItem(curMode, "charrangemax");
                        fileWriter.Paragraph(escapedItem("range", "string", range));
                        fileWriter.Paragraph(escapedItem("rof", "string", Item.DamageModeTagItem(curMode, "rof")));
                        fileWriter.Paragraph(escapedItem("shots", "string", Item.DamageModeTagItem(curMode, "shots")));
                        fileWriter.Paragraph(escapedItem("rcl", "number", Item.DamageModeTagItem(curMode, "rcl")));
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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportTraits(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var x = myCharacter.ItemsByName("Size Modifier", (int)TraitTypes.Attributes);
            var sm = "0";
            if(x.Count > 0)
            {
                GCATrait gCATrait = (GCATrait)x[1];
                sm = gCATrait.Score.ToString();
            }

            fileWriter.Paragraph("<traits>");
            fileWriter.Paragraph(escapedItem("race", "string", myCharacter.Race));
            fileWriter.Paragraph(escapedItem("height", "string", myCharacter.Height));
            fileWriter.Paragraph(escapedItem("weight", "string", myCharacter.Weight));
            fileWriter.Paragraph(escapedItem("age", "string", myCharacter.Age));
            fileWriter.Paragraph(escapedItem("appearance", "string", myCharacter.Appearance));
            fileWriter.Paragraph(escapedItem("sizemodifier", "string", sm));
            fileWriter.Paragraph(escapedItem("reach", "string", sm));
            fileWriter.Paragraph(escapedItem("tl", "string", myCharacter.TL));
            fileWriter.Paragraph(escapedItem("tl_points", "number", myCharacter.ItemByNameAndExt("ST", (int)TraitTypes.Stats).Points.ToString()));

            //Advantages
            exportAdvantages(myCharacter, fileWriter);
            //Disadvantages
            exportDisadvantages(myCharacter, fileWriter);
            //Cultural familiarities
            exportCuluralFamiliarty(myCharacter, fileWriter);
            //Languages
            exportLanguages(myCharacter, fileWriter);
            //reactionmodifiers
            exportReactionMods(myCharacter, fileWriter);
            fileWriter.Paragraph("</traits>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportAdvantages(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var i = 1;
            var Ads = myCharacter.ItemsByType[(int)TraitTypes.Advantages];
            var Templates = myCharacter.ItemsByType[(int)TraitTypes.Templates];
            var Perks = myCharacter.ItemsByType[(int)TraitTypes.Perks];
            var Features = myCharacter.ItemsByType[(int)TraitTypes.Features];

            fileWriter.Paragraph("<adslist>");

            foreach (GCATrait Template in Templates)
            {
                if (Template.Points >= 0)
                {
                    var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                    fileWriter.Paragraph(index);
                    fileWriter.Paragraph(escapedItem("name", "string", Template.DisplayName));
                    fileWriter.Paragraph(escapedItem("points", "number", Template.Points.ToString()));
                    fileWriter.Paragraph(escapedItem("text", "string", Template.Notes));
                    fileWriter.Paragraph(escapedItem("page", "string", Template.get_TagItem("page")));
                    fileWriter.Paragraph(index.Insert(1, "/"));
                    i++;
                }
                
            }

            foreach (GCATrait Adv in Ads)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(index);
                fileWriter.Paragraph(escapedItem("name", "string", Adv.DisplayName));
                fileWriter.Paragraph(escapedItem("points", "number", Adv.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", Adv.Notes));
                fileWriter.Paragraph(escapedItem("page", "string", Adv.get_TagItem("page")));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            foreach (GCATrait Perk in Perks)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(index);
                fileWriter.Paragraph(escapedItem("name", "string", Perk.DisplayName));
                fileWriter.Paragraph(escapedItem("points", "number", Perk.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", Perk.Notes));
                fileWriter.Paragraph(escapedItem("page", "string", Perk.get_TagItem("page")));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            foreach (GCATrait Feature in Features)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(index);
                fileWriter.Paragraph(escapedItem("name", "string", Feature.DisplayName));
                fileWriter.Paragraph(escapedItem("points", "number", Feature.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", Feature.Notes));
                fileWriter.Paragraph(escapedItem("page", "string", Feature.get_TagItem("page")));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            fileWriter.Paragraph("</adslist>");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportDisadvantages(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var i = 1;
            var Disads = myCharacter.ItemsByType[(int)TraitTypes.Disadvantages];
            var Templates = myCharacter.ItemsByType[(int)TraitTypes.Templates];
            var Quirks = myCharacter.ItemsByType[(int)TraitTypes.Quirks];

            fileWriter.Paragraph("<disadslist>");

            foreach (GCATrait Template in Templates)
            {
                if (Template.Points < 0)
                {
                    var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                    fileWriter.Paragraph(index);
                    fileWriter.Paragraph(escapedItem("name", "string", Template.DisplayName));
                    fileWriter.Paragraph(escapedItem("points", "number", Template.Points.ToString()));
                    fileWriter.Paragraph(escapedItem("text", "string", Template.Notes));
                    fileWriter.Paragraph(escapedItem("page", "string", Template.get_TagItem("page")));
                    fileWriter.Paragraph(index.Insert(1, "/"));
                    i++;
                }

            }

            foreach (GCATrait Disad in Disads)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(escapedItem("name", "string", Disad.DisplayName));
                fileWriter.Paragraph(escapedItem("points", "number", Disad.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", Disad.Notes));
                fileWriter.Paragraph(escapedItem("page", "string", Disad.get_TagItem("page")));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            foreach (GCATrait Quirk in Quirks)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(escapedItem("name", "string", Quirk.DisplayName));
                fileWriter.Paragraph(escapedItem("points", "number", Quirk.Points.ToString()));
                fileWriter.Paragraph(escapedItem("text", "string", Quirk.Notes));
                fileWriter.Paragraph(escapedItem("page", "string", Quirk.get_TagItem("page")));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }

            fileWriter.Paragraph("</disadslist>");

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportCuluralFamiliarty(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var i = 1;
            var Familirarities = myCharacter.ItemsByType[(int)TraitTypes.Cultures];

            fileWriter.Paragraph("<culturalfamiliaritylist>");
            foreach (GCATrait Familiarity in Familirarities)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(escapedItem("name", "string", Familiarity.DisplayName));
                fileWriter.Paragraph(escapedItem("points", "number", Familiarity.Points.ToString()));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }
            fileWriter.Paragraph("</culturalfamiliaritylist>");

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportLanguages(GCACharacter myCharacter, FileWriter fileWriter)
        {
            var i = 1;
            var Languages = myCharacter.ItemsByType[(int)TraitTypes.Languages];

            fileWriter.Paragraph("<languagelist>");
            foreach (GCATrait Language in Languages)
            {
                var index = "<id-" + i.ToString("D5", CultureInfo.CreateSpecificCulture("en-US")) + ">";
                fileWriter.Paragraph(escapedItem("name", "string", Language.DisplayName));
                fileWriter.Paragraph(escapedItem("spoken", "string", Language.LevelName));
                fileWriter.Paragraph(escapedItem("written", "string", Language.LevelName));
                fileWriter.Paragraph(escapedItem("points", "number", Language.Points.ToString()));
                fileWriter.Paragraph(index.Insert(1, "/"));
                i++;
            }
            fileWriter.Paragraph("</languagelist>");

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="myCharacter"></param>
        /// <param name="fileWriter"></param>
        private void exportReactionMods(GCACharacter myCharacter, FileWriter fileWriter)
        {
            GCATrait reaction = myCharacter.ItemByNameAndExt("Reaction", (int)TraitTypes.Attributes);
            string reactionmods = reaction.get_TagItem("bonuslist");
            reactionmods = reactionmods + ", " + reaction.get_TagItem("conditionallist");

            fileWriter.Paragraph(escapedItem("reactionmodifiers","string",reactionmods));
        }
        /// <summary>
        /// creates the xml tag with propper escaping of characters
        /// </summary>
        /// <param name="tagName"></param>
        /// <param name="tagType"></param>
        /// <param name="item"></param>
        /// <returns></returns>
        private string escapedItem(string tagName, string tagType, string item)
        {
            return "<" + tagName + " type=\"" + tagType + "\">"+  SecurityElement.Escape(item)  + "</" + tagName + ">";
        }

        /// <summary>
        /// utility function to check if a given item is hidden
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        private bool isItemHidden(GCATrait item)
        {
            return !(item.get_TagItem("hidden") == "");
        }
        /// <summary>
        /// assembles the damage string
        /// </summary>
        /// <param name="item"></param>
        /// <param name="mode"></param>
        /// <returns></returns>
        private string getDamageString(GCATrait item, int mode)
        {
            return item.DamageModeTagItem(mode, "chardamage") + " " + item.DamageModeTagItem(mode, "chardamtype");
        }

        private string getAdvantageName(GCATrait item)
        {
            var returnValue = "";
            returnValue = item.Name;

            if(item.NameExt.Length > 0)
            {
                returnValue = returnValue + " " + item.NameExt;
            }
                
            if (item.Level > 0)
            {
                var levelnames = item.LevelName.Split(',');

                if (levelnames.Length > 0)
                {
                    returnValue = returnValue + " (" + levelnames[item.Level] + ")";
                }
                else
                {
                    returnValue = returnValue + " (" + item.Level.ToString() + ")";
                }

                item.get_TagItem("initmods");
            }
            
            return returnValue;
        }
    }

}
