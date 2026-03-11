using ErwinAPI_AddIn;
using ErwinAPI_AddIn.Infra;
using ErwinAPI_AddIn.Models;
using SCAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace ErwinAPI_AddIn
{
    public class ConexaoErwin
    {

        //SESSÕES 
        public Session SessaoM0 { get; set; } //Sessão usada para alterar e criar objetos
        public Session SessaoM1 { get; set; } //Sessão usada para criar UDPs

        private PersistenceUnit ObjPU;
        private Application ObjApi;
        public string ConexaoArquivo;
        public string MartArquivo;

        public PersistenceUnit AbrirModeloArquivo(string arquivoErwin)
        {
            ConexaoArquivo = arquivoErwin;
            try
            {
                ObjApi = new Application();    
                //Recebe caminho do modelo no disco para add em PersistenceUnits 
                ObjPU = ObjApi.PersistenceUnits.Add(ConexaoArquivo);

            }

            catch (System.Runtime.InteropServices.COMException comEx)
            {
                if (comEx.Message.Contains("File exists and read only"))
                {
                    throw new Exception("Arquivo já aberto e definido como somente leitura!");
                }
                else
                {
                    throw new Exception("Erro na utilização da DLL do Erwin: " + comEx.Message);
                }
            }
            return ObjPU;
        }
        public List<PersistenceUnit> LerModelosAbertos()
        {
            List<PersistenceUnit> lstModelosAbertos = new List<PersistenceUnit>();
            try
            {  //trazer modelos de instancia aberta 
                Application application = new Application();
                foreach (PersistenceUnit persistenceUnit in application.PersistenceUnits)
                {
                    lstModelosAbertos.Add(persistenceUnit);
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                throw new Exception("Erro na utilização da DLL do Erwin: " + comEx.Message);
            }
            catch (Exception e)
            {
                throw new Exception("Erro na conexão com o Modelo de Dados: " + e.Message);
            }
            return lstModelosAbertos;
        }
        public void ConectaModeloArquivo(string arquivoErwin)
        {
            ConexaoArquivo = arquivoErwin;
            try
            {
                ObjApi = new SCAPI.Application();
                Boolean intRetOper;
                PropertyBag objPropertyBag = new SCAPI.PropertyBag();

                objPropertyBag.Add("Model_Type", "Combined");

                ObjPU = ObjApi.PersistenceUnits.Create(objPropertyBag);

                ObjPU = ObjApi.PersistenceUnits.Add(ConexaoArquivo, "RDO=no;");

                if (ObjApi.Sessions.Count > 0)
                    ObjApi.Sessions.Clear();

                SessaoM0 = ObjApi.Sessions.Add();
                SessaoM1 = ObjApi.Sessions.Add();
                intRetOper = SessaoM0.Open(ObjPU, SCAPI.SC_SessionLevel.SCD_SL_M0);
                intRetOper = SessaoM1.Open(ObjPU, SCAPI.SC_SessionLevel.SCD_SL_M1);

            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                throw new Exception("Erro na utilização da DLL do Erwin: " + comEx.Message);
            }
            catch (Exception e)
            {
                throw new Exception("Erro na conexão com o Modelo de Dados: " + e.Message);
            }
        }
        public void ConectaModeloMemoria(PersistenceUnit persistenceUnit)
        {
            try
            {
                ObjPU = persistenceUnit;
                // ISSUE: variable of a compiler-generated type
                Application application = new Application();
                // ISSUE: reference to a compiler-generated method
                SessaoM0 = application.Sessions.Add();
                // ISSUE: reference to a compiler-generated method
                SessaoM0.Open(ObjPU, SC_SessionLevel.SCD_SL_M0);
            }
            catch (Exception ex)
            {
                throw new Exception("Erro: conectaModeloErwin(objpu_):", ex);
            }
        }
        public void ConectaModeloMart(string martArquivo)
        {
            MartArquivo = martArquivo;
            //Mart://Mart/ErwinAPI/eMovies_Alterado?TRC=NO;SRV=192.168.30.1;PRT=18170;ASR=MartServer;UID=Administrator;PSW=;
            try
            {
                ObjApi = new SCAPI.Application();
                Boolean intRetOper;
                PropertyBag objPropertyBag = new SCAPI.PropertyBag();

                objPropertyBag.Add("Model_Type", "Combined");

                ObjPU = ObjApi.PersistenceUnits.Create(objPropertyBag);

                ObjPU = ObjApi.PersistenceUnits.Add(MartArquivo, "RDO=Yes;");

                if (ObjApi.Sessions.Count > 0)
                    ObjApi.Sessions.Clear();

                SessaoM0 = ObjApi.Sessions.Add();
                SessaoM1 = ObjApi.Sessions.Add();
                intRetOper = SessaoM0.Open(ObjPU, SCAPI.SC_SessionLevel.SCD_SL_M0);
                intRetOper = SessaoM1.Open(ObjPU, SCAPI.SC_SessionLevel.SCD_SL_M1);

            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                throw new Exception("Erro na utilização da DLL do Erwin: " + comEx.Message);
                //conexão ja foi aberta pelo Erwin. 
            }
            catch (Exception e)
            {
                throw new Exception("Erro na conexão com o Modelo de Dados: " + e.Message);
            }
        }
        public void AlteraPropriedade(ModelObject objeto, string nomeProp, string valorPropr)
        {
            var transIdAlter = SessaoM0.BeginNamedTransaction("ALTER PROP");
            objeto.Properties[nomeProp].Value = valorPropr;
            SessaoM0.CommitTransaction(transIdAlter);
        }
        public void GravaModelo(string origem)
        {
            if (origem == "file")
            {
                ObjPU.Save(ConexaoArquivo, "OVF=Yes");
            }
            if (origem == "memory")
            {
                ObjPU.Save(null, "OVF=Yes");
            }
            if (origem == "mart")
            {
                ObjPU.Save(MartArquivo, "RDO=NO;OVM=YES");
            }
        }
        public void DesconectaModelo()
        {
            try
            {
                if (SessaoM0 != null)
                    SessaoM0.Close();

                if (SessaoM1 != null)
                    SessaoM1.Close();

                if (ObjApi != null)
                    ObjApi.PersistenceUnits.Clear();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao fechar a conexão com o Modelo de Dados: " + ex.Message);
            }
            finally
            {
                if (ObjPU != null)
                    while (Marshal.ReleaseComObject(ObjPU) > 0) ;
                if (SessaoM0 != null)
                    while (Marshal.ReleaseComObject(SessaoM0) > 0) ;
                if (SessaoM1 != null)
                    while (Marshal.ReleaseComObject(SessaoM1) > 0) ;
                if (ObjApi != null)
                    while (Marshal.ReleaseComObject(ObjApi) > 0) ;
            }
        }
        public int getTotalModelosAbertos()
        {
            List<string> stringList = new List<string>();
            try
            {
                // ISSUE: variable of a compiler-generated type
                Application application = new Application();
                return application.PersistenceUnits.Count;
            }
            catch (Exception )
            {
                return -1;
            }
        }
        public PersistenceUnit LerModeloSelecionado(string modelo, string objectId)
        {
            List<PersistenceUnit> persistenceUnits = new List<PersistenceUnit>();
            persistenceUnits = LerModelosAbertos();
            foreach (PersistenceUnit pp in persistenceUnits)
            {
                if (modelo == pp.Name && objectId == pp.ObjectId)
                {
                    return pp;
                }
            }
            return null;
        }
        public List<ErwinTable> RetornaObjetoErwin(ConexaoErwin conexaoErwin)
        {
            try
            {
                List<ErwinTable> listErwinTable = new List<ErwinTable>();

                //Cria objeto com dados do Erwin para pegar os Entitades
                foreach (ModelObject modelObject in conexaoErwin.SessaoM0.ModelObjects)
                {
                    if (modelObject.ClassName.ToString() == "Entity")
                    {
                        listErwinTable.Add(new ErwinTable() { Name = modelObject.Name, ObjectType = modelObject.ClassName, Columns = new List<ErwinColumn>() });
                    }
                }

                //Cria objeto com dados do Erwin para pegar os Atributos e vincular as entidades
                foreach (ModelObject modelObject in conexaoErwin.SessaoM0.ModelObjects)
                {
                    if (modelObject.ClassName.ToString() == "Attribute")
                    {
                        ErwinTable erwinTable = listErwinTable.Where(c => c.Name == modelObject.Context.Name).FirstOrDefault();
                        if (erwinTable != null)
                        {
                            bool fkProp = false;

                            try
                            {
                                fkProp = modelObject.Properties["Parent_Attribute_Ref"] != null ? true : false;
                            }
                            catch (Exception)
                            {
                                fkProp = false;
                            }

                            erwinTable.Columns.Add(new ErwinColumn()
                            {
                                DataType = modelObject.Properties["Logical_Data_Type"].Value,
                                FK = fkProp,
                                PK = modelObject.Properties["Type"].Value == 0 ? true : false,
                                //Identity = bool.Parse(modelObject.Properties["Identity"].ToString()),
                                Name = modelObject.Name,
                                ParentDomainRef = modelObject.Properties["Parent_Domain_Ref"].Value
                            });
                        }
                    }
                }
                return listErwinTable;
            }
            catch (Exception )
            {
                return new List<ErwinTable>();
            }
        }      
        public void PadronizarNomeLogico_PK(ModelObjects modelObjects, string chave, Models.ParametersConfig parametersConfig)
        {
                     
            string entityName = "";
            string shortName = "";
            string keyGroupName = "";
            string keyGroupType = "";

            foreach (ModelObject modelObject in modelObjects)
            {
                //entity
                if (modelObject.ClassName.ToString() == "Entity")
                {
                    entityName = modelObject.Name;
                    shortName = modelObject.Properties.HasProperty(parametersConfig.udpShortNameLogical) == false ? "NO_ShortName" : modelObject.Properties[parametersConfig.udpShortNameLogical].Value;
                }
                //KeyGroup  //utilizando sequencia de leitura padrão do Erwin apos ler Entity
                if (modelObject.ClassName.ToString() == "Key_Group")
                {
                    keyGroupName = modelObject.Name;
                    if (modelObject.Properties["Key_Group_Type"].Value != null)
                    {
                        keyGroupType = modelObject.Properties["Key_Group_Type"].Value;
                       
                        if (keyGroupType.Contains(chave))
                        {                            
                            string valorNomenclaturaPK = entityName + "_" + shortName + "_" + keyGroupType;                        

                            var transIdAlterNomenclaturaPK = SessaoM0.BeginNamedTransaction("ALTER Nomenclatura_PK - Logical");
                            modelObject.Properties["Name"].Value = valorNomenclaturaPK;
                            //modelObject.Properties["Name"].Value = valorNomenclaturaPK;
                            SessaoM0.CommitTransaction(transIdAlterNomenclaturaPK);
                            GravaModelo("memory");                           
                        }
                    }
                }
            }
        }
        public void PadronizarNomeFisico_PK(ModelObjects modelObjects,string chave,Models.ParametersConfig parametersConfig)
        {
            string entityName = "";
            string shortName = "";
            string keyGroupName = "";
            string keyGroupType = "";
            foreach (ModelObject modelObject in modelObjects)
            {
                //entity
                if (modelObject.ClassName.ToString() == "Entity")
                {
                    entityName = modelObject.Name;
                    shortName = modelObject.Properties.HasProperty(parametersConfig.udpShortNamePhysical) == false ? "NO_ShortName" : modelObject.Properties[parametersConfig.udpShortNamePhysical].Value;
                }
                //KeyGroup  //utilizando sequencia de leitura padrão do Erwin apos ler Entity
                if (modelObject.ClassName.ToString() == "Key_Group")
                {
                    keyGroupName = modelObject.Name;
                    if (modelObject.Properties["Key_Group_Type"].Value != null)
                    {
                        keyGroupType = modelObject.Properties["Key_Group_Type"].Value;
                        if (keyGroupType.Contains(chave))
                        {
                            string valorNomenclaturaPK = entityName + "_" +  shortName + "_" + keyGroupType;
                            var transIdAlterNomenclaturaPK = SessaoM0.BeginNamedTransaction("ALTER Nomenclatura_PK - Physical");
                            //modelObject.Properties["Name"].Value = valorNomenclaturaPK;
                            modelObject.Properties["Physical_Name"].Value = valorNomenclaturaPK;
                            SessaoM0.CommitTransaction(transIdAlterNomenclaturaPK);
                            GravaModelo("memory");                           
                        }
                    }
                }
            }
        }
        public void PadronizarNomeLogico_FK(ModelObjects objetos, Models.ParametersConfig parametersConfig)
        {
            int count = 0;
            List<ModelObject> modelObjectList = new List<ModelObject>();
            List<string> listaFK = new List<string>();
            foreach (ModelObject modelObject in objetos)
            {
                if (modelObject.ClassName == "Relationship")
                {
                    modelObjectList.Add(modelObject);
                }
            }

            foreach (ModelObject model in modelObjectList)
            {
                var aa = model.Name;
                string relacObjectId = model.ObjectId;
                string relacName = model.Name;
                string parentEntityRef = "";
                string childEntityRef = "";
                string parentShortName = "";
                string childShortName = "";
                string parentEntityName = "";
                string childEntityName = "";
         
                //PARENT
                if (model.Properties["Parent_Entity_Ref"].Value != null)
                {
                    parentEntityRef = model.Properties["Parent_Entity_Ref"].Value;
                }
                foreach (ModelObject modelObjectParent in objetos)
                {
                    if (modelObjectParent.ClassName.ToString() == "Entity" && modelObjectParent.ObjectId == parentEntityRef)

                    {
                        parentShortName = modelObjectParent.Properties.HasProperty(parametersConfig.udpShortNameLogical) == false ? "NO_ShortName" : modelObjectParent.Properties[parametersConfig.udpShortNameLogical].Value;
                        parentEntityName = modelObjectParent.Name;
                    }
                }

                //CHILD
                if (model.Properties["Child_Entity_Ref"].Value != null)
                {
                    childEntityRef = model.Properties["Child_Entity_Ref"].Value;
                }
                
                foreach (ModelObject modelObjectChild in objetos)
                {
                    if (modelObjectChild.ClassName.ToString() == "Entity" && modelObjectChild.ObjectId == childEntityRef)
                    {
                        childShortName = modelObjectChild.Properties.HasProperty(parametersConfig.udpShortNameLogical) == false ? "NO_ShortName" : modelObjectChild.Properties[parametersConfig.udpShortNameLogical].Value;
                        childEntityName = modelObjectChild.Name;                      
                    }
                }

                string valorFK = childShortName + parentShortName;
                listaFK.Add(valorFK);

                count = listaFK.Where(s => s.Contains(valorFK)).Count();
                string valorNomenclaturaFK = childShortName + "_" + parentShortName + count + "_" + "FK";
                var transIdAlterNomenclaturaFK = SessaoM0.BeginNamedTransaction("ALTER Nomenclatura_FK");
                model.Properties["Name"].Value = valorNomenclaturaFK;
                model.Properties["Physical_Name"].Value = valorNomenclaturaFK;
                SessaoM0.CommitTransaction(transIdAlterNomenclaturaFK);

            }           
        }
        public void PadronizarNomeLogico_AK(ModelObjects modelObjects, string chave, Models.ParametersConfig parametersConfig)
        {
            string entityName = "";
            string shortName = "";
            string keyGroupName = "";
            string keyGroupType = "";

            foreach (ModelObject modelObject in modelObjects)
            {
                //entity
                if (modelObject.ClassName.ToString() == "Entity")
                {
                    entityName = modelObject.Name;
                    shortName = modelObject.Properties.HasProperty(parametersConfig.udpShortNameLogical) == false ? "NO_ShortName" : modelObject.Properties[parametersConfig.udpShortNameLogical].Value;
                }

                //KeyGroup  //utilizando sequencia de leitura padrão do Erwin apos ler Entity
                if (modelObject.ClassName.ToString() == "Key_Group")
                {
                    keyGroupName = modelObject.Name;

                    if (modelObject.Properties["Key_Group_Type"].Value != null)
                    {
                        keyGroupType = modelObject.Properties["Key_Group_Type"].Value;                    
                     
                        if (keyGroupType.Contains(chave))
                        {
                            string sequencial = Regex.Replace(keyGroupType, @"[^\d]", "");
                            //string valorNomenclaturaUK = shortName + "_" + keyGroupType;
                            string valorNomenclaturaUK = shortName + sequencial + "_" + "UK";

                            var transIdAlterNomenclaturaUK = SessaoM0.BeginNamedTransaction("ALTER Nomenclatura_UK");
                            modelObject.Properties["Name"].Value = valorNomenclaturaUK;
                            modelObject.Properties["Physical_Name"].Value = valorNomenclaturaUK;
                            SessaoM0.CommitTransaction(transIdAlterNomenclaturaUK);
                        }
                    }
                }
            }
        }
        public void PadronizarNomeLogico_IE(ModelObjects modelObjects, string chave, ParametersConfig parametersConfig)
        {
            string entityName = "";
            string shortName = "";
            string keyGroupName = "";
            string keyGroupType = "";
            foreach (ModelObject modelObject in modelObjects)
            {
                //entity
                if (modelObject.ClassName.ToString() == "Entity")
                {
                    entityName = modelObject.Name;
                    shortName = modelObject.Properties.HasProperty(parametersConfig.udpShortNameLogical) == false ? "NO_ShortName" : modelObject.Properties[parametersConfig.udpShortNameLogical].Value;
                }

                //KeyGroup  //utilizando sequencia de leitura padrão do Erwin apos ler Entity
                if (modelObject.ClassName.ToString() == "Key_Group")
                {
                    keyGroupName = modelObject.Name;

                    if (modelObject.Properties["Key_Group_Type"].Value != null)
                    {
                        keyGroupType = modelObject.Properties["Key_Group_Type"].Value;

                        if (keyGroupType.Contains(chave))
                        {
                            string sequencial = Regex.Replace(keyGroupType, @"[^\d]", "");                       
                            string valorNomenclaturaIX = shortName + sequencial + "_" + "IX";
                            var transIdAlterNomenclaturaUK = SessaoM0.BeginNamedTransaction("ALTER Nomenclatura_IX");
                            modelObject.Properties["Name"].Value = valorNomenclaturaIX;
                            modelObject.Properties["Physical_Name"].Value = valorNomenclaturaIX;
                            SessaoM0.CommitTransaction(transIdAlterNomenclaturaUK);
                        }
                    }
                }
            }
        }
        public void PadronizarNomeLogico_CK(ModelObjects modelObjects, ParametersConfig parametersConfig)
        {
            //Check
            foreach (ModelObject modelObject in modelObjects)
            {
                if (modelObject.ClassName.ToString() == "Check_Constraint_Usage")
                {
                    var id = modelObject.ObjectId.ToString();
                    var validationName = modelObject.Name;
                    var attributeName = modelObject.Context.Name;                   
                    var entityName = modelObject.Context.Context.Name;
                    var shortName = modelObject.Context.Context.Properties.HasProperty(parametersConfig.udpShortNameLogical) == false ? "NO_ShortName" : modelObject.Context.Context.Properties[parametersConfig.udpShortNameLogical].Value;                   
                    string valorNomenclaturaCK = shortName + "_" + attributeName + "_" + "CK";
                    var transIdAlterNomenclaturaCK = SessaoM0.BeginNamedTransaction("ALTER Nomenclatura_CK");
                    modelObject.Properties["Name"].Value = valorNomenclaturaCK;
                    modelObject.Properties["Physical_Name"].Value = valorNomenclaturaCK;
                    SessaoM0.CommitTransaction(transIdAlterNomenclaturaCK);
                }
            }
        }
        public ModelObject CriaObjeto(ModelObject objetoPai, string classe)
        {
            try
            {
                //Begin transaction
                var transIdCreate = SessaoM0.BeginNamedTransaction("Create Object");

                var objetosFilhos = SessaoM0.ModelObjects.Collect(objetoPai);
                var novoObjeto = objetosFilhos.Add(classe);
                SessaoM0.CommitTransaction(transIdCreate);
                return novoObjeto;
            }
            catch (Exception ex)
            {
                throw new Exception("Erro ao criar o objeto: " + ex.Message);
            }
        }
        public void ValidarDomain(ConexaoErwin conexaoErwin)
        {
            try
            {
                var listErwinDomain = new List<ErwinDomain>();
                foreach (ModelObject modelObject in conexaoErwin.SessaoM0.ModelObjects)
                {

                    if (modelObject.ClassName.ToString() == "Domain" && !(modelObject.Properties["Name"].Value.Contains("<")))
                    {
                        listErwinDomain.Add(new ErwinDomain
                        {
                            Long_Id = modelObject.Properties["Long_Id"].Value,
                            Name = modelObject.Properties["Name"].Value,
                            Logical_Data_Type = modelObject.Properties["Logical_Data_Type"] == null ? "" : modelObject.Properties["Logical_Data_Type"].Value,
                            Physical_Data_Type = modelObject.Properties["Physical_Data_Type"] == null ? "" : modelObject.Properties["Physical_Data_Type"].Value
                        });
                    }
                }

                foreach (ModelObject modelObject in conexaoErwin.SessaoM0.ModelObjects)
                {
                    Console.WriteLine("Objeto: " + modelObject.Name + " > Classe: " + modelObject.ClassName);

                    if (modelObject.ClassName.ToString() == "Attribute")
                    {
                        foreach (ErwinDomain item in listErwinDomain)
                        {
                            //Nomes iguais (Dominio x Atributos)
                            if (modelObject.Name == item.Name)
                            {
                                conexaoErwin.AlteraPropriedade(modelObject, "Parent_Domain_Ref", item.Long_Id);
                                conexaoErwin.AlteraPropriedade(modelObject, "Physical_Data_Type", item.Physical_Data_Type);
                                conexaoErwin.AlteraPropriedade(modelObject, "Logical_Data_Type", item.Logical_Data_Type);
                            }

                            //FL para FLAG
                            //Validar o FL somente para as duas primeiras letras
                            if (modelObject.Name.ToUpper().Contains("FL") && modelObject.Name != "FlAtivo" && item.Name == "FLAG")
                            {
                                conexaoErwin.AlteraPropriedade(modelObject, "Parent_Domain_Ref", item.Long_Id);
                                conexaoErwin.AlteraPropriedade(modelObject, "Physical_Data_Type", item.Physical_Data_Type);
                                conexaoErwin.AlteraPropriedade(modelObject, "Logical_Data_Type", item.Logical_Data_Type);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Erro na validação do Domain: " + ex.Message);
            }
        }
    }
}
