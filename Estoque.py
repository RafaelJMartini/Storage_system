from tkinter import *
from tkinter.ttk import Combobox
import psycopg2
import os
import shutil
import xml.etree.ElementTree as ET
import time
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import json
from datetime import datetime

with open("config.json") as f:
    config = json.load(f)


def lerxml():
    log = ""
    xmls_dir = '.\\xmls'
    xmls_dir_old = '.\\xmls_old'

    try:
        os.listdir(xmls_dir)
    except FileNotFoundError:
        os.makedirs(xmls_dir)

    try:
        os.listdir(xmls_dir_old)
    except FileNotFoundError:
        os.makedirs(xmls_dir_old)


    qtd_sucesso = 0
    qtd_falha = 0

    arquivos = [f for f in os.listdir(xmls_dir) if f.endswith('.xml')]
    pre = "{http://www.portalfiscal.inf.br/nfe}"
    for arquivo in arquivos:
        caminho_arq = os.path.join(xmls_dir,arquivo)
        print(f"\nlendo arquivo {arquivo}")

        try:
            tree = ET.parse(caminho_arq)
            root = tree.getroot()
            
            print(f"Raiz do XML: {root.tag}")

            for elem in root:
                if elem.attrib.get('versao'):
                    versao = elem.attrib.get('versao')
            
            if versao != '4.00':
                qtd_falha += 1
                log += f"\nERRO: A versão do arquivo {arquivo} é {versao}, o programa só aceita a versão 4.00"
                print(f"A versão do XML é {versao}, o programa só aceita a versão 4.00 (CHAMAR O RAFAEL)")


                continue

            #Busca o ID da nota
            protNFe = root.find(f'{pre}protNFe')
            infProt = protNFe.find(f'{pre}infProt')
            id_nf = infProt.find(f'{pre}chNFe').text

            print(f"ID da nota fiscal: {id_nf}")


            NFe = root.find(f'{pre}NFe')
            infNFe = NFe.find(f'{pre}infNFe')

            #busca dados da empresa emissora
            emit = infNFe.find(f'{pre}emit')
            CNPJ = emit.find(f'{pre}CNPJ').text
            empresaNome = emit.find(f'{pre}xNome').text
            NomeFantasia = emit.find(f'{pre}xFant').text
            print(f"Empresa {empresaNome} de CNPJ {CNPJ} com o nome fantasia {NomeFantasia}")

            #busca o valor da nota fiscal
            total = infNFe.find(f'{pre}total')
            ICMSTot = total.find(f'{pre}ICMSTot')
            valorNF = ICMSTot.find(f'{pre}vNF').text
            print(f"Valor da NF é {valorNF}")

            #busca a Data e hora da emissão da nota
            ide = infNFe.find(f'{pre}ide')
            datah = ide.find(f'{pre}dhEmi').text
            print(f"Data da emissão da nota fiscal: {datah}")

            produtos = []

            det = infNFe.findall(f'{pre}det')
            qtd_itens = len(det)
            for i,produto in enumerate(det):
                prod = produto.find(f'{pre}prod')
                cProd = prod.find(f'{pre}cProd').text
                xProd = prod.find(f'{pre}xProd').text
                qCom = prod.find(f'{pre}qCom').text
                ncm = prod.find(f'{pre}NCM').text
                vProd = prod.find(f'{pre}vProd').text
                produtos.append((cProd,xProd,qCom,ncm,vProd))
                
            try:
                conn = psycopg2.connect(**config)

                cursor = conn.cursor()
                
                #Verificar se a nota já está no banco
                query = f'''
                SELECT 1 FROM nfs WHERE id = %s

                '''
                cursor.execute(query,(id_nf,))

                resultado = cursor.fetchone()
                if resultado:
                    log+=f"\nA nota fiscal {arquivo} já foi inserida no sistema"
                    qtd_falha += 1
                    continue
                


                query_adic = f'''

                INSERT INTO fornecedores(
                cnpj,
                razao_social,
                nome_fantasia
                )
                VALUES(%s,%s,%s)
                ON CONFLICT (cnpj) DO NOTHING;

                INSERT INTO nfs(
                id,
                valortotal,
                cnpj,
                dhemi,
                qtd_itens            
                )
                VALUES(%s,%s,%s,%s,%s)
                ON CONFLICT (id) DO NOTHING;

                INSERT INTO adicao(
                datah,
                idnf
                )
                VALUES(%s,%s)
                
                RETURNING id;

                '''

                tempo_atual = time.localtime()
                datahora = time.strftime("%d/%m/%Y %H:%M:%S", tempo_atual)
                cursor.execute(query_adic,(CNPJ,empresaNome,NomeFantasia,id_nf,valorNF,CNPJ,datah,qtd_itens,datahora,id_nf))
                idadicao = cursor.fetchone()[0]

                

                print(produtos)
                for i in range(qtd_itens):
                    print(f"o I é {i}")

                    query_produtosAdic = f'''
                
                    

                    INSERT INTO produtosexternos(cnpj,codigoexterno)
                    VALUES(%s,%s)
                    ON CONFLICT (cnpj, codigoexterno) DO UPDATE 
                    SET codigoexterno = EXCLUDED.codigoexterno
                    RETURNING idproduto;

                    '''
                    cursor.execute(query_produtosAdic,(CNPJ,produtos[i][0]))
                    conn.commit()

                    idinternoproduto = cursor.fetchone()[0]
                    print(f"o id interno desse prod é {idinternoproduto}")

                    query_insertproduto = f'''

                    INSERT INTO produtos(
                    idproduto,
                    nomeprod,
                    ncm,
                    quant
                    )
                    VALUES(%s,%s,%s,%s)
                    ON CONFLICT (idproduto) DO UPDATE 
                    SET quant = produtos.quant + EXCLUDED.quant;



                    INSERT INTO produtosdaadicao(idadicao,idproduto,quant)
                    VALUES(%s,%s,%s);
                    
                    

                    ''' 

                    cursor.execute(query_insertproduto,(idinternoproduto,produtos[i][1],produtos[i][3],produtos[i][2],idadicao,idinternoproduto,produtos[i][2]))
                    conn.commit()
                    
                        
                    #cursor.execute(query_produtos,(produtos[i][1],produtos[i][2],produtos[i][3]))
                    #conn.commit()

                print("Produtos adicionados com sucesso!")

                print("Dados adicionados com sucesso!")


            except psycopg2.Error as e:
                log += "\nFALHA AO CONECTAR AO BANCO"
                print(f"Erro ao conectar no banco:{str(e)}")
            
            finally:
                if 'conn' in locals() and conn:
                    cursor.close()
                    conn.close()
                    print("Conexão fechada")

            


        except ET.ParseError as e:
            print(f"Erro ao processar o arquivo{arquivo}: {e}")
        log += f"\n{arquivo} OK"
        qtd_sucesso+=1
        #Move os arquivos para o xml_old (DESCOMENTAR QUANDO ESTIVER TUDO PRONTO)
        shutil.move(xmls_dir+'\\'+arquivo, xmls_dir_old+'\\'+arquivo)
        print(f"movido {arquivo} de {xmls_dir} para {xmls_dir_old}")


    
    if qtd_falha:
        log += f"\n{qtd_falha} XMLs não foram inseridos por conta de um erro!"
    if qtd_sucesso:
        log += f"\n{qtd_sucesso} XMLs adicionados com sucesso"
    msg["text"] = log


def gerar_excel():
    log = ""
    excel_dir = ".\\excel"
    try:
        os.listdir(excel_dir)
    except FileNotFoundError:
        os.makedirs(excel_dir)
    excel_dir = ".\\excel\\EstoqueProdutos.xlsx"

    try:
        conn = psycopg2.connect(**config)

        query = '''
        SELECT idproduto AS ID,nomeprod AS Nome_do_Produto,quant AS Quantidade
        FROM produtos


        '''
        
        cursor = conn.cursor()
        cursor.execute(query)
        dados_excel = cursor.fetchall()

        colunas = [desc[0] for desc in cursor.description]  
        df = pd.DataFrame(dados_excel, columns=colunas)
        df.columns = ['ID','Nome do Produto','Quantidade']
        
        df.to_excel(excel_dir,index=False)

        wb = load_workbook(excel_dir)
        ws = wb.active

        for col in ws.columns:
            max_length = 0
            column = col[0].column_letter  # Pega a letra da coluna
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)  # Adicionando um espaço extra
            ws.column_dimensions[column].width = adjusted_width



        # Salvando as alterações no Excel
        wb.save(excel_dir) 



        cursor.close()
        conn.close()

    except psycopg2.Error as e:
        log += "\nFALHA AO CONECTAR AO BANCO"
        print(f"Erro ao conectar no banco:{str(e)}")
        


    log += "\nExcel gerado com sucesso"
    msg["text"] = log

def add_rm_prod():
    produtos_dict = {}
    keys = {}
    def consulta_banco():
        try:
            conn = psycopg2.connect(**config)

            cursor = conn.cursor()
                    
            #Verificar se a nota já está no banco
            query = f'''
            SELECT idproduto,nomeprod,quant FROM produtos
            '''
            cursor.execute(query)
            
            nonlocal produtos_dict
            nonlocal keys
            #quantidades = []
            # Iterando pelos resultados da consulta e criando o dicionário
            for row in cursor.fetchall():
                idproduto, nomeprod, quant = row
                chave = f"{idproduto} - {nomeprod}"  # Você pode usar o idproduto + nomeprod como chave
                produtos_dict[chave] = quant  # Quantidade é o valor
                keys[chave] = idproduto
                #quantidades.append(quant)
            cursor.close()
            conn.close()

        except psycopg2.Error as e:
            log = "\nFALHA AO CONECTAR AO BANCO"
            print(f"Erro ao conectar no banco:{str(e)}")
            return log
    consulta_banco()


    def formatar_quantidade(valor):
        """ Arredonda a quantidade para o máximo possível (inteiro se não tiver decimal) """
        valor = float(valor)
        valor = round(valor, 2)  # Arredonda para 2 casas
        return int(valor) if valor == int(valor) else valor  # Remove ".0" se for inteiro

    def atualizar_quantidade(*_):


        """ Atualiza o texto com a quantidade do item selecionado """
        item_selecionado = btn_selecionaprod.get()  # Pega o item selecionado no combobox
        quantidade = produtos_dict.get(item_selecionado, "Indisponível")  # Busca a quantidade no dicionário
        try:
            quantidade = float(quantidade)
        except ValueError:
            return None
        if opcao_var.get() == "Remover":
            
            maximoitem = "(Max: {})".format(formatar_quantidade(quantidade))
            txtmax['text'] = maximoitem
            txtprodinvalid['text'] = ''
            
        else:
            txtmax['text'] = ""
        return quantidade
        

    def escolha_rm():
        opcao_var.set("Remover")
        print("Escolha atual:", opcao_var.get())
        btn.config(text=f"{opcao_var.get()}")
        atualizar_quantidade()
    
    def escolha_add():
        opcao_var.set("Adicionar")
        print("Escolha atual:", opcao_var.get())
        btn.config(text=f"{opcao_var.get()}")
        txtmax['text'] = ""

    log = ""
    janela_add_rm = Tk()
    janela_add_rm.title("Adicionar ou remover um produto")
    janela_add_rm.geometry('600x400+650+200')

    #------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    
    opcao_var = StringVar()

    radio1 = Radiobutton(janela_add_rm, text="Adicionar", variable=opcao_var, value="Adicionar", command=escolha_add)
    radio2 = Radiobutton(janela_add_rm, text="Remover", variable=opcao_var, value="Remover", command=escolha_rm)

    opcao_var.set("Adicionar")

    # Posiciona os botões na janela
    radio1.pack(pady=10)
    radio2.pack()
    



    lista_produtos = keys.keys()
    lista_produtos = list(lista_produtos)
    btn_selecionaprod = Combobox(janela_add_rm,values=lista_produtos)
    btn_selecionaprod.pack(pady=40,ipadx=150)
    btn_selecionaprod.set("Escolha um produto")


#------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
 
    def ao_clicar(event):
        if entrada.get() == "Digite a quantidade":
            entrada.delete(0, "end")  # Remove o placeholder
            entrada.config(fg="black")  # Muda a cor do texto

    def ao_sair(event):
        if not entrada.get():
            entrada.insert(0, "Digite a quantidade")  # Reinsere o placeholder
            entrada.config(fg="gray")  # Volta a cor para cinza
    entrada = Entry(janela_add_rm)
    entrada.pack(pady=10,ipadx=50)
    entrada.insert(0, "Digite a quantidade")  # Define o placeholder

    # Eventos para limpar/recuperar o placeholder
    entrada.bind("<FocusIn>", ao_clicar)
    entrada.bind("<FocusOut>", ao_sair)

    txtmax = Label(janela_add_rm,text="")
    txtmax.pack()

    txtprodinvalid = Label(janela_add_rm,text="")
    txtprodinvalid.pack()

    btn_selecionaprod.bind("<<ComboboxSelected>>", atualizar_quantidade)

    def remover_adicionar():
        nonlocal log
        if btn_selecionaprod.get() not in lista_produtos:
            txtprodinvalid["text"] = "Por favor, selecione um produto válido."
            return
        try:
            numero = float(entrada.get())  # Pega o valor digitado no campo
        except ValueError:
            txtprodinvalid['text'] = 'Por favor, insira um número válido'
            return

        quantidade = atualizar_quantidade(None)
        if quantidade == None:
            txtprodinvalid["text"] = "Por favor, insira um número válido."
            return
        if opcao_var.get() == "Remover":
            if numero > quantidade:
                txtprodinvalid["text"] = "Por favor, insira um número válido."
                return
        
        try:
            numero = int(numero)  # Converte para inteiro (ou float se quiser)
            print(f"Número inserido: {numero}")
        except ValueError:
            txtprodinvalid["text"] = "Por favor, insira um número válido."
            return
        if numero <= 0:
            txtprodinvalid['text'] = 'Por favor, insira um número válido'

    
        try:

            conn = psycopg2.connect(**config)

            cursor = conn.cursor()
            boleanadic = 0
            if opcao_var.get() == "Remover":
                nova_quant = quantidade - numero
                boleanadic = 1
            else:
                nova_quant = quantidade + numero
                boleanadic = 0

            query = f'''
            INSERT INTO historicoestoque(
            datah,
            idproduto,
            quant,
            adicao_remocao
            )
            VALUES(%s,%s,%s,%s)
            '''
            datah = datetime.now()
            cursor.execute(query,(datah,int(keys[btn_selecionaprod.get()]),numero,boleanadic))

            query = f'''
            UPDATE produtos
            SET quant = %s
            WHERE idproduto = %s
            '''
            
            cursor.execute(query,(nova_quant,int(keys[btn_selecionaprod.get()])))

            if opcao_var.get() == "Remover":
                print(f"Adicionando {numero} de {quantidade} do produto de ID {int(keys[btn_selecionaprod.get()])}")
            else:
                print(f"Adcionando {numero} de {quantidade} do produto de ID {int(keys[btn_selecionaprod.get()])}")
            conn.commit()


            cursor.close()
            conn.close()
            if opcao_var.get() == "Remover":
                log = "\nProdutos removidos com sucesso"
            else:
                log = "\nProdutos adicionados com sucesso"
            consulta_banco()
            quantidade = atualizar_quantidade(None)
            txtprodinvalid["text"] += log

        except psycopg2.Error as e:
            log += f"\nFALHA AO CONECTAR AO BANCO PARA ADICIONAR OU REMOVER O PRODUTO: {str(e)}"
            print(f"Erro ao conectar no banco:{str(e)}")

    btn = Button(janela_add_rm, text=opcao_var.get(), command=remover_adicionar)
    btn.pack()

    janela_add_rm.mainloop()




#window config
win = Tk()
win.title("Estoque JC Ferreira")
win.geometry("800x600+500+100")
win.configure(bg="White")
win.iconbitmap(".\\content\\icone-storage.ico")

titulo = Label(win,text="Sistema de Estoque JC Ferreira",font=("Arial",20),bg="white")
titulo.pack(pady=(10,60))


#Buttons
btnlerxml = Button(win,text="Ler XMLs",font=("Arial",12),command=lerxml)
btnlerxml.pack(pady=(0,15),ipadx=80)

btngerarexcel = Button(win,text="Gerar Excel",font=("Arial",12),command=gerar_excel)
btngerarexcel.pack(pady=(0,15),ipadx=80)

btnremoverprod = Button(win,text="Adicionar/Remover Produto",font=("Arial",12),command=add_rm_prod)
btnremoverprod.pack(pady=(0,15),ipadx=80)

msg = Label(win,text=" ",bg="white")
msg.pack(pady=5)


win.mainloop()
