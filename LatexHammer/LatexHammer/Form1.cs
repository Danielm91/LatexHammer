using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace LatexHammer
{
   
    public partial class Form1 : Form
    {
        public const int MODEL_TYPE_INFANTRY = 0;
        public const int MODEL_TYPE_JUMP_INFANTRY = 1;
        public const int MODEL_TYPE_JET_PACK_INFANTRY = 2;
        public const int MODEL_TYPE_MONSTROUS_CREATURE = 3;
        public const int MODEL_TYPE_JUMP_MONSTROUS_CREATURE = 4;
        public const int MODEL_TYPE_FLYING_MONSTROUS_CREATURE =5;
        public const int MODEL_TYPE_JET_PACK_MONSTROUS_CREATURE = 6;
        public const int MODEL_TYPE_VEHICLE = 7;
        public const int MODEL_TYPE_FLYER = 8;
        public const int MODEL_TYPE_WALKER = 9;
        public const int MODEL_TYPE_BEAST = 10;

        public string[] MODEL_TYPE_STRINGS = new string[11] {"Infantry", "Jump Infantry", "Jet Pack Infantry",
            "Monstrous Creature", "Jump Monstrous Creature", "Flying Monstrous Creature", "Jet Pack Monstrous Creature",
            "Vehicle", "Flyer", "Walker", "Beast"};

        public struct model
        {
            public string name;
            public int type;
            public int ws;
            public int bs;
            public int s;
            public int t;
            public int w;
            public int i;
            public int a;
            public int ld;
            public int sv;
            public int inv;

            //Vehicles only
            public int fr;
            public int si;
            public int re;
            public int hp;
        }

        public struct wargear
        {
            public string name;
            public string description;
        }

        public struct special
        {
            public string name;
            public string description;
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            currentRuleSet.masterModelList = new List<model> { };
            currentRuleSet.masterWargearList = new List<wargear> { };
            currentRuleSet.masterSpecialList = new List<special> { };

            currentArmyList.listOfDetachments = new List<detachment> { };
    }

        public struct ruleSet
        {
            public List<model> masterModelList;
            public List<wargear> masterWargearList;
            public List<special> masterSpecialList;
        }

        public ruleSet currentRuleSet;

        public void refreshRuleSet()
        {
            lbModels.Items.Clear();
            lbWargear.Items.Clear();
            lbSpecialRules.Items.Clear();
            cbUnitModelType.Items.Clear();
            cbUnitSpecialType.Items.Clear();
            cbUnitWargearType.Items.Clear();

            for(int index = 0; index < currentRuleSet.masterModelList.Count; index ++)
            {
                lbModels.Items.Add(currentRuleSet.masterModelList[index].name);
                cbUnitModelType.Items.Add(currentRuleSet.masterModelList[index].name);
            }

            for (int index = 0; index < currentRuleSet.masterWargearList.Count; index++)
            {
                lbWargear.Items.Add(currentRuleSet.masterWargearList[index].name);
                cbUnitWargearType.Items.Add(currentRuleSet.masterWargearList[index].name);
            }

            for (int index = 0; index < currentRuleSet.masterSpecialList.Count; index++)
            {
                lbSpecialRules.Items.Add(currentRuleSet.masterSpecialList[index].name);
                cbUnitSpecialType.Items.Add(currentRuleSet.masterSpecialList[index].name);
            }
        }

        private void btnModelAdd_Click(object sender, EventArgs e)
        {
            model newModel;
            newModel.name = txtModelName.Text;
            newModel.type = cbModelType.SelectedIndex;
            newModel.ws = int.Parse(txtModelWS.Text);
            newModel.bs = int.Parse(txtModelBs.Text);
            newModel.s = int.Parse(txtModelS.Text);
            newModel.t = int.Parse(txtModelT.Text);
            newModel.w = int.Parse(txtModelW.Text);
            newModel.i = int.Parse(txtModelI.Text);
            newModel.a = int.Parse(txtModelA.Text);
            newModel.ld = int.Parse(txtModelLd.Text);
            newModel.sv = int.Parse(txtModelSave.Text);
            newModel.inv = int.Parse(txtModelInv.Text);
            newModel.fr = int.Parse(txtModelFr.Text);
            newModel.si = int.Parse(txtModelSi.Text);
            newModel.re = int.Parse(txtModelRe.Text);
            newModel.hp = int.Parse(txtModelHp.Text);
            lbModels.Items.Add(newModel.name);
            cbUnitModelType.Items.Add(newModel.name);
            currentRuleSet.masterModelList.Add(newModel);
        }

        private void txtModelSave_Click(object sender, EventArgs e)
        {
            if (lbModels.SelectedIndex >= 0)
            {
                model newModel;
                newModel.name = txtModelName.Text;
                newModel.type = cbModelType.SelectedIndex;
                newModel.ws = int.Parse(txtModelWS.Text);
                newModel.bs = int.Parse(txtModelBs.Text);
                newModel.s = int.Parse(txtModelS.Text);
                newModel.t = int.Parse(txtModelT.Text);
                newModel.w = int.Parse(txtModelW.Text);
                newModel.i = int.Parse(txtModelI.Text);
                newModel.a = int.Parse(txtModelA.Text);
                newModel.ld = int.Parse(txtModelLd.Text);
                newModel.sv = int.Parse(txtModelSave.Text);
                newModel.inv = int.Parse(txtModelInv.Text);
                newModel.fr = int.Parse(txtModelFr.Text);
                newModel.si = int.Parse(txtModelSi.Text);
                newModel.re = int.Parse(txtModelRe.Text);
                newModel.hp = int.Parse(txtModelHp.Text);
                lbModels.Items[lbModels.SelectedIndex] =newModel.name;
                cbUnitModelType.Items[lbModels.SelectedIndex] = newModel.name;
                currentRuleSet.masterModelList[lbModels.SelectedIndex] = newModel;
            }
        }

        private void btnModelRemove_Click(object sender, EventArgs e)
        {
            if (lbModels.SelectedIndex >= 0)
            {
                currentRuleSet.masterModelList.RemoveAt(lbModels.SelectedIndex);
                cbUnitModelType.Items.RemoveAt(lbModels.SelectedIndex);
                lbModels.Items.RemoveAt(lbModels.SelectedIndex);
            }
        }

        private void lbModels_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbModels.SelectedIndex >= 0)
            {
                txtModelName.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].name;
                cbModelType.SelectedIndex = currentRuleSet.masterModelList[lbModels.SelectedIndex].type;
                txtModelWS.Text  = currentRuleSet.masterModelList[lbModels.SelectedIndex].ws.ToString();
                txtModelBs.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].bs.ToString();
                txtModelS.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].s.ToString();
                txtModelT.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].t.ToString();
                txtModelW.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].w.ToString();
                txtModelI.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].i.ToString();
                txtModelA.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].a.ToString();
                txtModelLd.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].ld.ToString();
                txtModelSave.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].sv.ToString();
                txtModelInv.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].inv.ToString();
                txtModelFr.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].fr.ToString();
                txtModelSi.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].si.ToString();
                txtModelRe.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].re.ToString();
                txtModelHp.Text = currentRuleSet.masterModelList[lbModels.SelectedIndex].hp.ToString();
            }
        }

        private void btnWargearAdd_Click(object sender, EventArgs e)
        {
            wargear newWargear;
            newWargear.name = txtWargearName.Text;
            newWargear.description = txtWargearDescription.Text;
            lbWargear.Items.Add(newWargear.name);
            cbUnitWargearType.Items.Add(newWargear.name);
            currentRuleSet.masterWargearList.Add(newWargear);       
        }

        private void btnWargearSave_Click(object sender, EventArgs e)
        {
            if (lbWargear.SelectedIndex >= 0)
            {
                wargear newWargear;
                newWargear.name = txtWargearName.Text;
                newWargear.description = txtWargearDescription.Text;
                lbWargear.Items[lbWargear.SelectedIndex] = newWargear.name;
                cbUnitWargearType.Items[lbWargear.SelectedIndex] = newWargear.name;
                currentRuleSet.masterWargearList[lbWargear.SelectedIndex] = newWargear;
            }
        }

        private void btnWargearRemove_Click(object sender, EventArgs e)
        {
            if (lbWargear.SelectedIndex >= 0)
            {
                currentRuleSet.masterWargearList.RemoveAt(lbWargear.SelectedIndex);
                cbUnitWargearType.Items.RemoveAt(lbWargear.SelectedIndex);
                lbWargear.Items.RemoveAt(lbWargear.SelectedIndex);
            }
    }

        private void lbWargear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbWargear.SelectedIndex >= 0)
            {
                txtWargearName.Text = currentRuleSet.masterWargearList[lbWargear.SelectedIndex].name;
                txtWargearDescription.Text = currentRuleSet.masterWargearList[lbWargear.SelectedIndex].description;
            }
        }

        private void btnSpecialAdd_Click(object sender, EventArgs e)
        {
            special newSpecial;
            newSpecial.name = txtSpecialRuleName.Text;
            newSpecial.description = rtbSpecialRulesDescription.Text;
            currentRuleSet.masterSpecialList.Add(newSpecial);
            lbSpecialRules.Items.Add(newSpecial.name);
            cbUnitSpecialType.Items.Add(newSpecial.name);
        }

        private void btnSpecialSave_Click(object sender, EventArgs e)
        {
            if (lbSpecialRules.SelectedIndex >= 0)
            {
                special newSpecial;
                newSpecial.name = txtSpecialRuleName.Text;
                newSpecial.description = rtbSpecialRulesDescription.Text;
                currentRuleSet.masterSpecialList[lbSpecialRules.SelectedIndex] = newSpecial;
                lbSpecialRules.Items[lbSpecialRules.SelectedIndex] = newSpecial.name;
                cbUnitSpecialType.Items[lbSpecialRules.SelectedIndex] = newSpecial.name;
            }
        }

        private void btnSpecialRemove_Click(object sender, EventArgs e)
        {
            if (lbSpecialRules.SelectedIndex >= 0)
            {
                currentRuleSet.masterSpecialList.RemoveAt(lbSpecialRules.SelectedIndex);
                cbUnitSpecialType.Items.RemoveAt(lbSpecialRules.SelectedIndex);
                lbSpecialRules.Items.RemoveAt(lbSpecialRules.SelectedIndex);
            }
        }

        private void lbSpecialRules_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbSpecialRules.SelectedIndex >= 0)
            {
                txtSpecialRuleName.Text = currentRuleSet.masterSpecialList[lbSpecialRules.SelectedIndex].name;
                rtbSpecialRulesDescription.Text = currentRuleSet.masterSpecialList[lbSpecialRules.SelectedIndex].description;
            }
        }

        //Army stuff
        public struct detachment
        {
            public string name;
            public List<unit> listOfUnits;

        }

        public int calcDetachmentPoints(detachment input)
        {
            int returnValue = 0;

            for(int index = 0; index < input.listOfUnits.Count; index ++)
            {
                returnValue += calcUnitPoints(input.listOfUnits[index]);
            }

            return returnValue;
        }

        public struct armyList
        {
            public List<detachment> listOfDetachments;
        }

        public int calcArmyListPoints(armyList input)
        {
            int returnValue = 0;

            for (int index = 0; index < input.listOfDetachments.Count; index++)
            {
                returnValue += calcDetachmentPoints(input.listOfDetachments[index]);
            }

            return returnValue;
        }

        public armyList currentArmyList;

        public void refreshArmyList()
        {
            lbDetachment.Items.Clear();
            lbUnits.Items.Clear();
            lbUnitModels.Items.Clear();
            lbUnitWargear.Items.Clear();
            lbUnitSpecials.Items.Clear();

            for(int index = 0; index < currentArmyList.listOfDetachments.Count; index ++)
            {
                lbDetachment.Items.Add(currentArmyList.listOfDetachments[index].name);
            }
        }
        private void btnDetachmentAdd_Click(object sender, EventArgs e)
        {
            detachment newDetachment;
            newDetachment.name = txtDetachmentName.Text;

            newDetachment.listOfUnits = new List<unit> { };

            currentArmyList.listOfDetachments.Add(newDetachment);
            lbDetachment.Items.Add(newDetachment.name);
        }

        private void btnDetachmentSave_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0)
            {
                detachment bufferDetachment = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex];
                bufferDetachment.name = txtDetachmentName.Text;

                lbDetachment.Items[lbDetachment.SelectedIndex] = bufferDetachment.name;
                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex] = bufferDetachment;
            }
        }

        private void btnDetachmentRemove_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0)
            {
                currentArmyList.listOfDetachments.RemoveAt(lbDetachment.SelectedIndex);
                lbDetachment.Items.RemoveAt(lbDetachment.SelectedIndex);
            }
        }

        private void lbDetachment_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0)
            {
                txtDetachmentName.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].name;

                lbUnits.Items.Clear();
                lbUnitModels.Items.Clear();
                lbUnitWargear.Items.Clear();
                lbUnitSpecials.Items.Clear();

                for (int index = 0; index < currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits.Count; index ++)
                {
                    lbUnits.Items.Add(currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[index].name);
                }

                txtDetachmentPoints.Text = calcDetachmentPoints(currentArmyList.listOfDetachments[lbDetachment.SelectedIndex]).ToString();
            }
        }

        //Stuff for actual units
        public struct modelAspect
        {
            public int type;
            public int qty;
            public int pts;
            public List<unitAspect> listOfWargear;
            public List<unitAspect> listOfSpecials;
        }

        public struct unitAspect
        {
            public int type;
            public int qty;
            public int pts;
        }


        public struct unit
        {
            public string name;
            public int role;

            public List<modelAspect> listOfModels;
            
        }

        //Unit Role Constants
        public const int UNIT_ROLE_HQ = 0;
        public const int UNIT_ROLE_ELITES = 1;
        public const int UNIT_ROLE_TROOPS = 2;
        public const int UNIT_ROLE_FAST_ATTACK = 3;
        public const int UNIT_ROLE_HEAVY_SUPPORT = 4;
        public const int UNIT_ROLE_LORD_OF_WAR = 5;
        public const int UNIT_ROLE_FORTIFICATION = 6;

        public void updatePointDisplayOnUnitChange()
        {
            txtUnitPoints.Text = calcUnitPoints(currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex]).ToString();
            txtDetachmentPoints.Text = calcDetachmentPoints(currentArmyList.listOfDetachments[lbDetachment.SelectedIndex]).ToString();
            txtArmyPoints.Text = calcArmyListPoints(currentArmyList).ToString();
        }

        public int calcUnitPoints(unit input)
        {
            int returnValue = 0;

            for(int index = 0; index < input.listOfModels.Count; index ++)
            {
                returnValue += input.listOfModels[index].qty * input.listOfModels[index].pts;
                for(int wargearIndex = 0; wargearIndex < input.listOfModels[index].listOfWargear.Count; wargearIndex ++)
                {
                    returnValue += input.listOfModels[index].listOfWargear[wargearIndex].qty * input.listOfModels[index].listOfWargear[wargearIndex].pts;
                }
                for(int specialIndex = 0; specialIndex < input.listOfModels[index].listOfSpecials.Count; specialIndex ++)
                {
                    returnValue += input.listOfModels[index].listOfSpecials[specialIndex].pts;
                }
            }
            return returnValue;
        }

        private void txtUnitAdd_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0)
            {
                unit newUnit;
                newUnit.name = txtUnitName.Text;
                newUnit.role = cbUnitRole.SelectedIndex;

                newUnit.listOfModels = new List<modelAspect> { };
              
                lbUnits.Items.Add(newUnit.name);
                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits.Add(newUnit);
                lbUnitModels.Items.Clear();
                lbUnitWargear.Items.Clear();
                lbUnitSpecials.Items.Clear();
            }
            
        }

        private void txtUnitSave_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0)
            {
                unit bufferUnit = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex];
                bufferUnit.name = txtUnitName.Text;
                bufferUnit.role = cbUnitRole.SelectedIndex;
                lbUnits.Items[lbUnits.SelectedIndex] = bufferUnit.name;

                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex] = bufferUnit;
            }
        }

        private void txtUnitRemove_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0)
            {
                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits.RemoveAt(lbUnits.SelectedIndex);
                lbUnits.Items.RemoveAt(lbUnits.SelectedIndex);
                lbUnitModels.Items.Clear();
                lbUnitWargear.Items.Clear();
                lbUnitSpecials.Items.Clear();
            }
        }

        private void lbUnits_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0)
            {
                txtUnitName.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].name;
                cbUnitRole.SelectedIndex = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].role;

                lbUnitModels.Items.Clear();
                lbUnitWargear.Items.Clear();
                lbUnitSpecials.Items.Clear();

                for(int index = 0; index < currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels.Count; index ++)
                {
                    lbUnitModels.Items.Add(currentRuleSet.masterModelList[currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[index].type].name);
                }

                updatePointDisplayOnUnitChange();
            }
        }

        private void btnUnitModelAdd_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >=0 & cbUnitModelType.SelectedIndex >= 0)
            {
                modelAspect newModel;
                newModel.type = cbUnitModelType.SelectedIndex;
                newModel.qty = int.Parse(txtUnitModelQty.Text);
                newModel.pts = int.Parse(txtUnitModelPts.Text);
                if (newModel.qty > 0)
                {
                    newModel.listOfWargear = new List<unitAspect> { };
                    newModel.listOfSpecials = new List<unitAspect> { };

                    lbUnitModels.Items.Add(cbUnitModelType.Items[cbUnitModelType.SelectedIndex]);
                    currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels.Add(newModel);
                }

                updatePointDisplayOnUnitChange();
            }
        }

        private void btnUnitModelSave_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & cbUnitModelType.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0)
            {
                modelAspect newModel = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex];
                newModel.type = cbUnitModelType.SelectedIndex;
                newModel.qty = int.Parse(txtUnitModelQty.Text);
                newModel.pts = int.Parse(txtUnitModelPts.Text);
                if (newModel.qty > 0)
                {
                    lbUnitModels.Items[lbUnitModels.SelectedIndex] = cbUnitModelType.Items[cbUnitModelType.SelectedIndex];
                    currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex] = newModel;
                }
                else
                {
                    currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels.RemoveAt(lbUnitModels.SelectedIndex);
                    lbUnitModels.Items.RemoveAt(lbUnitModels.SelectedIndex);
                }

                updatePointDisplayOnUnitChange();
            }
        }

        private void lbUnitModels_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0)
            {
                cbUnitModelType.SelectedIndex = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].type;
                txtUnitModelQty.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].qty.ToString();
                txtUnitModelPts.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].pts.ToString();

                lbUnitWargear.Items.Clear();
                lbUnitSpecials.Items.Clear();

                for (int index = 0; index < currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear.Count; index ++)
                {
                    lbUnitWargear.Items.Add(currentRuleSet.masterWargearList[currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear[index].type].name);
                }

                for (int index = 0; index < currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials.Count; index++)
                {
                    lbUnitSpecials.Items.Add(currentRuleSet.masterSpecialList[currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials[index].type].name);
                }
            }
        }

        private void btnUnitWargearAdd_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & cbUnitWargearType.SelectedIndex >= 0)
            {
                unitAspect newWargear;
                newWargear.type = cbUnitWargearType.SelectedIndex;
                newWargear.pts = int.Parse(txtUnitWargearPts.Text);
                newWargear.qty = int.Parse(txtUnitWargearQty.Text);

                if (newWargear.qty > 0)
                {
                    currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear.Add(newWargear);
                    lbUnitWargear.Items.Add(currentRuleSet.masterWargearList[newWargear.type].name);
                }

                updatePointDisplayOnUnitChange();
            }
        }

        private void btnUnitWargearSave_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & cbUnitWargearType.SelectedIndex >= 0 & lbUnitWargear.SelectedIndex >= 0)
            {
                unitAspect newWargear;
                newWargear.type = cbUnitWargearType.SelectedIndex;
                newWargear.pts = int.Parse(txtUnitWargearPts.Text);
                newWargear.qty = int.Parse(txtUnitWargearQty.Text);

                if (newWargear.qty > 0)
                {
                    currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear[lbUnitWargear.SelectedIndex] = newWargear;
                    lbUnitWargear.Items[lbUnitWargear.SelectedIndex] = currentRuleSet.masterWargearList[newWargear.type].name;
                }
                else
                {
                    currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear.RemoveAt(lbUnitWargear.SelectedIndex);
                    lbUnitWargear.Items.RemoveAt(lbUnitWargear.SelectedIndex);
                }

                updatePointDisplayOnUnitChange();
            }
        }

        private void lbUnitWargear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & lbUnitWargear.SelectedIndex >= 0)
            {
                cbUnitWargearType.SelectedIndex = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear[lbUnitWargear.SelectedIndex].type;
                txtUnitWargearPts.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear[lbUnitWargear.SelectedIndex].pts.ToString();
                txtUnitWargearQty.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfWargear[lbUnitWargear.SelectedIndex].qty.ToString();
            }

        }

        private void btnUnitSpecialAdd_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & cbUnitSpecialType.SelectedIndex >= 0)
            {
                unitAspect newSpecial;
                newSpecial.type = cbUnitSpecialType.SelectedIndex;
                newSpecial.pts = int.Parse(txtUnitSpecialPts.Text);
                newSpecial.qty = 0;

                lbUnitSpecials.Items.Add(currentRuleSet.masterSpecialList[cbUnitSpecialType.SelectedIndex].name);
                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials.Add(newSpecial);

                updatePointDisplayOnUnitChange();
            }
        }

        private void btnUnitSpecialSave_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & cbUnitSpecialType.SelectedIndex >= 0 & lbUnitSpecials.SelectedIndex >= 0)
            {
                unitAspect newSpecial;
                newSpecial.type = cbUnitSpecialType.SelectedIndex;
                newSpecial.pts = int.Parse(txtUnitSpecialPts.Text);
                newSpecial.qty = 0;

                lbUnitSpecials.Items[lbUnitSpecials.SelectedIndex] = currentRuleSet.masterSpecialList[cbUnitSpecialType.SelectedIndex].name;
                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials[lbUnitSpecials.SelectedIndex] = newSpecial;

                updatePointDisplayOnUnitChange();
            }
        }

        private void btnUnitSpecialRemove_Click(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & lbUnitSpecials.SelectedIndex >= 0)
            {         
                currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials.RemoveAt(lbUnitSpecials.SelectedIndex);
                lbUnitSpecials.Items.RemoveAt(lbUnitSpecials.SelectedIndex);

                updatePointDisplayOnUnitChange();
            }
        }

        private void lbUnitSpecials_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbDetachment.SelectedIndex >= 0 & lbUnits.SelectedIndex >= 0 & lbUnitModels.SelectedIndex >= 0 & lbUnitSpecials.SelectedIndex >= 0)
            {
                cbUnitSpecialType.SelectedIndex = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials[lbUnitSpecials.SelectedIndex].type;
                txtUnitSpecialPts.Text = currentArmyList.listOfDetachments[lbDetachment.SelectedIndex].listOfUnits[lbUnits.SelectedIndex].listOfModels[lbUnitModels.SelectedIndex].listOfSpecials[lbUnitSpecials.SelectedIndex].pts.ToString();             
            }
       }

        private void btnMasterListExport_Click(object sender, EventArgs e)
        {
            sfdRuleList.ShowDialog();
        }

        private void sfdRuleList_FileOk(object sender, CancelEventArgs e)
        {
            string ruleSetString = Newtonsoft.Json.JsonConvert.SerializeObject(currentRuleSet);
            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(sfdRuleList.FileName))
            {
                file.Write(ruleSetString);
            }
        }

        private void btnMasterListImport_Click(object sender, EventArgs e)
        {
            ofdRuleList.ShowDialog();
        }

        private void ofdRuleList_FileOk(object sender, CancelEventArgs e)
        {
            string ruleSetString;
            using (System.IO.StreamReader file =
           new System.IO.StreamReader(ofdRuleList.FileName))
            {
                ruleSetString = file.ReadToEnd();
                currentRuleSet = Newtonsoft.Json.JsonConvert.DeserializeObject<ruleSet>(ruleSetString);
            }

            refreshRuleSet();
        }

        private void btnArmyListExport_Click(object sender, EventArgs e)
        {
            sfdArmyList.ShowDialog();
        }

        private void sfdArmyList_FileOk(object sender, CancelEventArgs e)
        {
            string armyListString = Newtonsoft.Json.JsonConvert.SerializeObject(currentArmyList);
            using (System.IO.StreamWriter file =
           new System.IO.StreamWriter(sfdArmyList.FileName))
            {
                file.Write(armyListString);
            }
        }

        private void btnArmyListImport_Click(object sender, EventArgs e)
        {
            ofdArmyList.ShowDialog();
        }

        private void ofdArmyList_FileOk(object sender, CancelEventArgs e)
        {
            string armyListString;
            using (System.IO.StreamReader file =
           new System.IO.StreamReader(ofdArmyList.FileName))
            {
                armyListString = file.ReadToEnd();
                currentArmyList= Newtonsoft.Json.JsonConvert.DeserializeObject<armyList>(armyListString);
            }

            refreshArmyList();
        }

        private void btnGenerateTex_Click(object sender, EventArgs e)
        {
            sfdGenerateTex.ShowDialog();
        }

        private void sfdGenerateTex_FileOk(object sender, CancelEventArgs e)
        {
            using (System.IO.StreamWriter file =
          new System.IO.StreamWriter(sfdGenerateTex.FileName))
            {
                file.WriteLine("\\documentclass[11pt]{article} % use larger type, default would be 10pt");
 
                file.WriteLine("");
                file.WriteLine("\\usepackage[utf8]{inputenc} % set input encoding (not needed with XeLaTeX)");

                file.WriteLine("");
                file.WriteLine("%%% PAGE DIMENSIONS");
                file.WriteLine("\\usepackage{geometry} % to change the page dimensions");
                file.WriteLine("\\geometry{a4paper} % or letterpaper(US) or a5paper or....");
                file.WriteLine("\\geometry{margin = 0.5in} % for example, change the margins to 2 inches all round");


                file.WriteLine("");
                file.WriteLine("\\usepackage{graphicx} % support the \\includegraphics command and options");

                file.WriteLine("");
                file.WriteLine("% \\usepackage[parfill]{parskip} % Activate to begin paragraphs with an empty line rather than an indent");

                file.WriteLine("");
                file.WriteLine("%%% PACKAGES");
                file.WriteLine("\\usepackage{booktabs} % for much better looking tables");
                file.WriteLine("\\usepackage{array} % for better arrays (eg matrices) in maths");
                file.WriteLine("\\usepackage{ragged2e}");
                file.WriteLine("\\usepackage{paralist} % very flexible & customisable lists(eg.enumerate / itemize, etc.)");
                file.WriteLine("\\usepackage{verbatim} % adds environment for commenting out blocks of text & for better verbatim");
                file.WriteLine("\\usepackage{subfig} % make it possible to include more than one captioned figure / table in a single float");


                file.WriteLine("%%% SECTION TITLE APPEARANCE");
                file.WriteLine("\\usepackage{sectsty}");
                file.WriteLine("\\allsectionsfont{\\sffamily\\mdseries\\upshape} % (See the fntguide.pdf for font help)");
                file.WriteLine("% (This matches ConTeXt defaults)");

                file.WriteLine("");
                file.WriteLine("%%% ToC(table of contents) APPEARANCE");
                file.WriteLine("\\usepackage[nottoc,notlof,notlot]{tocbibind} % Put the bibliography in the ToC");
                file.WriteLine("\\usepackage[titles,subfigure]{tocloft} % Alter the style of the Table of Contents");
                file.WriteLine("\\renewcommand{\\cftsecfont}{\\rmfamily\\mdseries\\upshape}");
                file.WriteLine("\\renewcommand{\\cftsecpagefont}{\\rmfamily\\mdseries\\upshape} % No bold!");

                file.WriteLine("");
                file.WriteLine("\\addtolength{\\topmargin}{-60pt}");
                file.WriteLine("\\addtolength{\\textheight}{60pt}");

                file.WriteLine("");
                file.WriteLine("%%% END Article customizations");

                file.WriteLine("");
                file.WriteLine("%%% The \"real\" document content comes below...");

                file.WriteLine("");
                file.WriteLine("\\title{" + txtArmyName.Text + " - " + calcArmyListPoints(currentArmyList) + "}");
                file.WriteLine("\\author{" + txtAuthor.Text + "}");
                file.WriteLine("\\date{ } % Activate to display a given date or no date (if empty),");
                file.WriteLine("% otherwise the current date is printed");

                file.WriteLine("\\renewcommand\\thesection{ }");
                file.WriteLine("\\renewcommand\\thesubsection{ }");
                file.WriteLine("\\renewcommand\\thesubsubsection{ }");

                file.WriteLine("");
                file.WriteLine("\\begin{document}");
                file.WriteLine("\\maketitle");

                file.WriteLine("\\begin{flushleft}");

                bool firstUnitFound = false;
                for (int indexDetachment = 0; indexDetachment < currentArmyList.listOfDetachments.Count; indexDetachment++)
                {
                    file.WriteLine("");
                    file.WriteLine("\\section{" + currentArmyList.listOfDetachments[indexDetachment].name + "}");

                    firstUnitFound = false;
                    for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {
                       

                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_HQ)
                        {
                            if (firstUnitFound == false)
                            {
                                firstUnitFound = true;
                                file.WriteLine("");
                                file.WriteLine("\\subsection{HQ}");
                            }

                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }
                        
                    }

                    firstUnitFound = false;
                    for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {

                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_ELITES)
                        {
                            if (firstUnitFound == false)
                            {
                                firstUnitFound = true;
                                file.WriteLine("");
                                file.WriteLine("\\subsection{Elites}");
                            }

                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }

                    }

                firstUnitFound = false;
                for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {
                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_TROOPS)
                        {
                        if (firstUnitFound == false)
                        {
                            firstUnitFound = true;
                            file.WriteLine("");
                            file.WriteLine("\\subsection{Troops}");
                        }

                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }

                    }

                    firstUnitFound = false;
                    for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {
                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_FAST_ATTACK)
                        {
                            if (firstUnitFound == false)
                            {
                                firstUnitFound = true;
                                file.WriteLine("");
                                file.WriteLine("\\subsection{Fast Attack}");
                            }
                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }

                    }

                    firstUnitFound = false;
                    for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {
                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_HEAVY_SUPPORT)
                        {
                            if (firstUnitFound == false)
                            {
                                firstUnitFound = true;
                                file.WriteLine("");
                                file.WriteLine("\\subsection{Heavy Support}");
                            }
                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }

                    }

                    firstUnitFound = false;
                    for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {
                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_LORD_OF_WAR)
                        {
                            if (firstUnitFound == false)
                            {
                                firstUnitFound = true;
                                file.WriteLine("");
                                file.WriteLine("\\subsection{Lord Of War}");
                            }
                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }

                    }

                    firstUnitFound = false;
                    for (int indexUnit = 0; indexUnit < currentArmyList.listOfDetachments[indexDetachment].listOfUnits.Count; indexUnit++)
                    {
                        if (currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].role == UNIT_ROLE_FORTIFICATION)
                        {
                            if (firstUnitFound == false)
                            {
                                firstUnitFound = true;
                                file.WriteLine("");
                                file.WriteLine("\\subsection{Fortification}");
                            }
                            file.WriteLine(writeUnitStats(indexDetachment, indexUnit));
                        }

                    }
                }
                                          
                file.WriteLine("");        
                file.WriteLine("\\end{flushleft}");
                file.WriteLine("\\end{document}");

            }
        }

        public string writeUnitStats(int indexDetachment, int indexUnit)
        {
            string returnString = "";

            modelAspect bufferModelAspect;
            bool firstRegularModel = false;
            bool firstVehicle = false;
            bool firstWalker = false;

            returnString += "\n";           
            returnString += "\\subsubsection{" + currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].name + " - " + calcUnitPoints(currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit]).ToString() + "}\n";

            for (int indexModel = 0; indexModel < currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels.Count; indexModel++)
            {
                bufferModelAspect = currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels[indexModel];
                if (currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_INFANTRY || currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_JUMP_INFANTRY || currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_JET_PACK_INFANTRY ||
                    currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_BEAST || currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_MONSTROUS_CREATURE || currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_FLYING_MONSTROUS_CREATURE ||
                    currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_JUMP_MONSTROUS_CREATURE)
                {
                    if (firstRegularModel == false)
                    {
                        firstRegularModel = true;

                        returnString += "\n";
                        returnString += "\\renewcommand{\\arraystretch}\n";
                        returnString += "{1}\n";
                        returnString += "\\begin{tabular}\n";
                        returnString += "{ | p{4cm} l *{9} {c}  >{\\raggedleft\\arraybackslash} p{4cm} |}\n";
                        returnString += "\\hline\n";
                        returnString += "& \\textbf{Qty}\n";
                        returnString += "& \\textbf{WS}\n";
                        returnString += "& \\textbf{BS}\n";
                        returnString += "& \\textbf{S}\n";
                        returnString += "& \\textbf{T}\n";
                        returnString += "& \\textbf{W}\n";
                        returnString += "& \\textbf{I}\n";
                        returnString += "& \\textbf{A}\n";
                        returnString += "& \\textbf{Ld}\n";
                        returnString += "& \\textbf{Sv}\n";
                        returnString += "& \\textbf{Unit Type} \\\\\n";
                       
                    }

                    returnString += currentRuleSet.masterModelList[bufferModelAspect.type].name + " & " + bufferModelAspect.qty.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].ws.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].bs.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].s.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].t.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].w.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].i.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].a.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].ld.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].sv.ToString() + "+\\textbackslash " + currentRuleSet.masterModelList[bufferModelAspect.type].inv.ToString() + "++ \n";
                    returnString += "& " + MODEL_TYPE_STRINGS[currentRuleSet.masterModelList[bufferModelAspect.type].type] + " \\\\ \n";

                }
               
            }
            
            if (firstRegularModel == true)
            {
                returnString += "\\hline\n"; //This one is for the full section conclusion
                returnString += "\\end{tabular}\n";
                returnString += "\\vspace{2mm}\n";
            }

            for (int indexModel = 0; indexModel < currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels.Count; indexModel++)
            {
                bufferModelAspect = currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels[indexModel];
                if (currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_VEHICLE || currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_FLYER)
                {
                    if (firstVehicle == false)
                    {
                        firstVehicle = true;
                        returnString += "\n";
                        returnString += "\\renewcommand{\\arraystretch}{1}\n";
                        returnString += "\\begin{tabular}{ | p{4cm}  l*{4}{c}  >{\\raggedleft\\arraybackslash}p{8.7cm} |}\n";
                        returnString += "\\hline\n";
                        returnString += " & \\textbf{Qty} & \\textbf{BS} &  \\textbf{FR} & \\textbf{SI} & \\textbf{RE}  & \\textbf{Unit Type} \\\\\n";
                       
                    }

                    returnString += currentRuleSet.masterModelList[bufferModelAspect.type].name + " & " + bufferModelAspect.qty.ToString() + "\n";                  
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].bs.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].fr.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].si.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].re.ToString() + "\n";                  
                    returnString += "& " + MODEL_TYPE_STRINGS[currentRuleSet.masterModelList[bufferModelAspect.type].type] + " \\\\ \n";
                }
            }
           

            if (firstVehicle == true)
            {
                returnString += "\\hline\n"; //This one is for the full section conclusion
                returnString += "\\end{tabular}\n";
                returnString += "\\vspace{2mm}\n";
            }

            for (int indexModel = 0; indexModel < currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels.Count; indexModel++)
            {
                bufferModelAspect = currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels[indexModel];
                if (currentRuleSet.masterModelList[bufferModelAspect.type].type == MODEL_TYPE_WALKER)
                {
                    if (firstWalker == false)
                    {
                        firstWalker = true;
                        returnString += "\n";
                        returnString += "\\renewcommand{\arraystretch}{1}\n";
                        returnString += "\\begin{tabular}{ | p{4cm}  l*{8}{c}  >{\\raggedleft\\arraybackslash}p{5.5cm} |}\n";
                        returnString += "\\hline\n";
                        returnString += "Name & \\textbf{Qty} & \\textbf{WS} & \\textbf{BS} &  \\textbf{FR} & \\textbf{SI} & \\textbf{RE} & \\textbf{S} & \\textbf{I} & \\textbf{A} & \\textbf{Unit Type} \\\\\n";
                       
                    }
                    returnString += currentRuleSet.masterModelList[bufferModelAspect.type].name + " & " + bufferModelAspect.qty.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].ws.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].bs.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].fr.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].si.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].re.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].s.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].i.ToString() + "\n";
                    returnString += "& " + currentRuleSet.masterModelList[bufferModelAspect.type].a.ToString() + "\n";
                    returnString += "& " + MODEL_TYPE_STRINGS[currentRuleSet.masterModelList[bufferModelAspect.type].type] + " \\\\ \n";
                }
            }

           

            if (firstWalker == true)
            {
                returnString += "\\hline\n"; //This one is for the full section conclusion
                returnString += "\\end{tabular}\n";
                returnString += "\\vspace{ 2mm}\n";
            }

           
            returnString += "\n";

           

            for (int indexModel = 0; indexModel < currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels.Count; indexModel++)
            {
                bufferModelAspect = currentArmyList.listOfDetachments[indexDetachment].listOfUnits[indexUnit].listOfModels[indexModel];
              
               
                returnString += "\\begin{tabular}\n";
                returnString += "{>{\\raggedright\\arraybackslash} p{6cm} p{10cm} }\n";
                returnString += "\\textbf{Wargear (" + currentRuleSet.masterModelList[bufferModelAspect.type].name + ")}\n";
                returnString += "& \\textbf{Special Rules (" + currentRuleSet.masterModelList[bufferModelAspect.type].name + ")}\\\\\n";
                returnString += "\\begin{itemize}[$\\bullet$]\n";
                returnString += "\\itemsep0em\n";
                if (bufferModelAspect.listOfWargear.Count > 0)
                {
                    for (int indexWargear = 0; indexWargear < bufferModelAspect.listOfWargear.Count; indexWargear++)
                    {
                        returnString += "\\item " + currentRuleSet.masterWargearList[bufferModelAspect.listOfWargear[indexWargear].type].name + "(" + currentRuleSet.masterWargearList[bufferModelAspect.listOfWargear[indexWargear].type].description + ")\n";
                    }
                }
                else
                {
                    returnString += "\\item none";
                }
                returnString += "\\end{itemize}\n";
                returnString += "&\n";
                returnString += "\\begin{itemize}\n";
                returnString += "[$\\bullet$]\n";
                if (bufferModelAspect.listOfSpecials.Count > 0)
                {
                    for (int indexSpecials = 0; indexSpecials < bufferModelAspect.listOfSpecials.Count; indexSpecials++)
                    {
                        returnString += "\\item " + currentRuleSet.masterSpecialList[bufferModelAspect.listOfSpecials[indexSpecials].type].name + " - " + currentRuleSet.masterSpecialList[bufferModelAspect.listOfSpecials[indexSpecials].type].description + "\n";
                    }
                }
                else
                {
                    returnString += "\\item none";
                }
                returnString += "\\end{itemize}\n";
                returnString += "\\end{tabular}\n";
            }

                
               
            
            return returnString;
        }
    }
}
