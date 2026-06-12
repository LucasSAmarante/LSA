import streamlit as st
import pandas as pd
from fpdf import FPDF
import datetime
import os
import re

# --- INICIALIZAÇÃO DA LISTA DE ITENS (SESSION STATE) ---
if 'lista_itens_orcamento' not in st.session_state:
    st.session_state.lista_itens_orcamento = []

# --- CLASSE DO PDF PREMIUM COM DESIGN DINÂMICO ---
class PDF_Assistencia_Design(FPDF):
    def header(self):
        # Cores de Identidade da L.S.A
        self.c_principal = (15, 23, 42)    # Slate Escuro (Sofisticado)
        self.c_destaque = (0, 150, 214)    # Azul Técnico / Cyan
        self.c_cinza_fundo = (248, 250, 252) # Fundo de blocos
        
        # Barra decorativa superior para quebrar a folha totalmente branca
        self.set_fill_color(*self.c_destaque)
        self.rect(0, 0, 210, 4, 'F')
        
        # Logo lateral ajustada
        if os.path.exists("logo.png"):
            self.image("logo.png", 12, 10, 28)
            self.set_x(45)
        else:
            self.set_x(12)
            
        # Textos do Cabeçalho institucional
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*self.c_principal)
        self.cell(0, 6, "L.S.A INFORMÁTICA", border=0, ln=1, align="L")
        
        self.set_font("Helvetica", "", 9)
        self.set_text_color(100, 116, 139)
        if os.path.exists("logo.png"): self.set_x(45)
        else: self.set_x(12)
        self.cell(0, 4, "Manutenção Especializada de Computadores e Notebooks", border=0, ln=1, align="L")
        
        if os.path.exists("logo.png"): self.set_x(45)
        else: self.set_x(12)
        self.cell(0, 4, "WhatsApp: (47) 9 9733-1906  |  lsainfo.informatica@gmail.com", border=0, ln=1, align="L")
        
        # Espaçamento inferior estável
        self.ln(14)

    def criar_painel_dados(self, dados):
        """ Cria um card fechado com fundo colorido para os dados de atendimento """
        x_init = 10
        y_init = self.get_y()
        
        # Define altura dinâmica baseada na existência ou não de diagnóstico longo
        altura_bloco = 38 if dados['diagnostico'] else 26
        
        # Desenha o fundo e contorno do card
        self.set_fill_color(248, 250, 252)
        self.set_draw_color(226, 232, 240)
        self.set_line_width(0.3)
        self.rect(x_init, y_init, 190, altura_bloco, 'DF')
        
        # Preenchimento interna das informações (Row 1)
        self.set_y(y_init + 3)
        self.set_x(x_init + 4)
        self.set_font("Helvetica", "B", 8.5)
        self.set_text_color(100, 116, 139)
        self.cell(95, 4, "CLIENTE / CONTATO", ln=0)
        self.cell(85, 4, "EQUIPAMENTO / MODELO", ln=1)
        
        self.set_x(x_init + 4)
        self.set_font("Helvetica", "", 10)
        self.set_text_color(15, 23, 42)
        self.cell(95, 5, f"{dados['cliente']} - {dados['contato']}", ln=0)
        self.cell(85, 5, dados['modelo'], ln=1)
        self.ln(2)
        
        # Row 2
        self.set_x(x_init + 4)
        self.set_font("Helvetica", "B", 8.5)
        self.set_text_color(100, 116, 139)
        self.cell(95, 4, "TÉCNICO RESPONSÁVEL", ln=0)
        self.cell(85, 4, "DATA SOLICITAÇÃO", ln=1)
        
        self.set_x(x_init + 4)
        self.set_font("Helvetica", "", 10)
        self.set_text_color(15, 23, 42)
        self.cell(95, 5, dados['tecnico'], ln=0)
        self.cell(85, 5, dados['data'], ln=1)
        
        # Row 3 (Diagnóstico em itálico destacado)
        if dados['diagnostico']:
            self.ln(2)
            self.set_x(x_init + 4)
            self.set_font("Helvetica", "B", 8.5)
            self.set_text_color(100, 116, 139)
            self.cell(0, 4, "DIAGNÓSTICO / OBSERVAÇÕES INICIAIS", ln=1)
            
            self.set_x(x_init + 4)
            self.set_font("Helvetica", "I", 9.5)
            self.set_text_color(51, 65, 85)
            self.cell(0, 5, dados['diagnostico'], ln=1)
            
        self.set_y(y_init + altura_bloco + 6)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(148, 163, 184)
        # Substituído '•' por '|' para evitar o erro de codificação Unicode
        self.cell(0, 10, f"L.S.A Informática  |  Navegantes - SC  |  Página {self.page_no()}", 0, 0, "C")

# --- FUNÇÕES DE TRANSPORTE DE DADOS ---
def extrair_preco(valor):
    try:
        return float(valor)
    except:
        numeros = re.findall(r'\d+', str(valor))
        if numeros:
            return float(numeros[0])
        return 0.0

@st.cache_data
def carregar_dados_excel():
    arquivo = "ORDEM DE SERVIÇO.xlsx"
    if os.path.exists(arquivo):
        try:
            df_c = pd.read_excel(arquivo, sheet_name="Clientes")
            df_s = pd.read_excel(arquivo, sheet_name="Serviços")
            if 'Código' in df_s.columns:
                df_s = df_s[df_s['Código'].notna() & (df_s['Código'].astype(str).str.startswith(('S', 'I'), na=False))]
            return df_c, df_s
        except:
            pass
    return None, None

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="L.S.A - Orçamentos Premium", page_icon="💻", layout="centered")

df_clientes, df_servicos = carregar_dados_excel()

st.title("📊 Painel de Orçamentos Dinâmicos - L.S.A")
st.write("Layout focado na experiência visual e clareza de dados comerciais.")

# 1. IDENTIFICAÇÃO DO ATENDIMENTO
st.subheader("👤 1. Identificação de Atendimento")

nome_cliente, tel_cliente = "", ""
if df_clientes is not None:
    lista_clientes = ["-- Selecionar Cliente Cadastrado --"] + df_clientes['Nome'].dropna().tolist()
    cliente_sel = st.selectbox("Buscar Cliente no Banco", lista_clientes)
    if cliente_sel != "-- Selecionar Cliente Cadastrado --":
        dados_c = df_clientes[df_clientes['Nome'] == cliente_sel].iloc[0]
        nome_cliente = str(dados_c['Nome'])
        tel_cliente = str(dados_c.get('Telefone', ''))

col1, col2 = st.columns(2)
with col1:
    cliente = st.text_input("Nome do Cliente", value=nome_cliente)
    modelo_equipamento = st.text_input("Modelo do Aparelho", placeholder="Ex: Notebook Samsung NP350XAA")
    data_solicitacao = st.date_input("Data de Entrada", datetime.date.today())
with col2:
    telefone = st.text_input("Celular / Contato", value=tel_cliente)
    tecnico_responsavel = st.text_input("Técnico Responsável", value="Lucas Souza do Amarante")

diagnostico = st.text_input("Diagnóstico Resumido (Exibe no painel do PDF)", placeholder="Ex: Equipamento sofreu queda e quebrou a tela.")

st.markdown("---")

# 2. TABELA DINÂMICA DE ITENS
st.subheader("🛠️ 2. Especificação de Peças e Mão de Obra")

desc_servico, preco_sugerido = "", 0.0
if df_servicos is not None:
    lista_servicos = ["-- Selecionar Serviço/Peça do Banco --"] + df_servicos['Serviço'].dropna().tolist()
    servico_sel = st.selectbox("Filtrar Tabela de Preços", lista_servicos)
    if servico_sel != "-- Selecionar Serviço/Peça do Banco --":
        dados_s = df_servicos[df_servicos['Serviço'] == servico_sel].iloc[0]
        desc_servico = str(dados_s['Serviço'])
        preco_sugerido = extrair_preco(dados_s.get('Valor (R$)', 0))

col_it1, col_it2, col_it3, col_it4 = st.columns([2, 0.5, 1, 1])
with col_it1:
    item_desc = st.text_input("Item ou Serviço", value=desc_servico)
with col_it2:
    item_qtd = st.number_input("Qtd", min_value=1, value=1, step=1)
with col_it3:
    item_base = st.number_input("Valor Normal (R$)", min_value=0.0, value=preco_sugerido, format="%.2f")
with col_it4:
    sugestao_pix = item_base * 0.93 if item_base > 0 else 0.0
    item_avista = st.number_input("Valor c/ Desc. PIX (R$)", min_value=0.0, value=sugestao_pix, format="%.2f")

if st.button("➕ Inserir Item na Tabela", use_container_width=True):
    if item_desc:
        st.session_state.lista_itens_orcamento.append({
            "item": item_desc,
            "qtde": item_qtd,
            "base": item_base * item_qtd,
            "avista": item_avista * item_qtd
        })
        st.toast("Item adicionado!")
    else:
        st.error("Preencha a descrição antes de adicionar.")

if st.session_state.lista_itens_orcamento:
    df_preview = pd.DataFrame(st.session_state.lista_itens_orcamento)
    df_preview.columns = ["Item / Serviço Técnico", "Qtd", "Total Base (R$)", "Total À Vista (R$)"]
    st.dataframe(df_preview, use_container_width=True, hide_index=True)
    
    if st.button("🗑️ Limpar Todos os Itens", type="secondary"):
        st.session_state.lista_itens_orcamento = []
        st.rerun()

st.markdown("---")

# 3. GARANTIA
st.subheader("📜 3. Termos Adicionais")
termos_garantia = st.text_area(
    "Cláusulas de Garantia Opcionais", 
    value="- 90 dias para defeitos relacionados exclusivamente ao serviço executado.\n- Quebras, quedas, oxidação, mau uso ou violação dos lacres internos do aparelho anulam a garantia automaticamente.\n- Componentes e peças de reposição possuem prazos balizados direto com o fabricante."
)

# COMPILAÇÃO E DOWNLOAD DO PDF ATRAENTE
if st.button("💥 GERAR ORÇAMENTO DINÂMICO PREMIUM", use_container_width=True, type="primary"):
    if not cliente:
        st.error("❌ O nome do cliente é obrigatório.")
    elif not st.session_state.lista_itens_orcamento:
        st.error("❌ Adicione pelo menos um item à tabela para gerar.")
    else:
        pdf = PDF_Assistencia_Design()
        pdf.add_page()
        
        # Renderiza o Card de Atendimento Inicial
        dados_atendimento = {
            "cliente": cliente,
            "modelo": modelo_equipamento if modelo_equipamento else "Não Informado",
            "contato": telefone if telefone else "Não Informado",
            "tecnico": tecnico_responsavel,
            "data": data_solicitacao.strftime('%d/%m/%Y'),
            "diagnostico": diagnostico
        }
        pdf.criar_painel_dados(dados_atendimento)
        
        # --- TABELA DE ITENS COM ALTA CARGA VISUAL ---
        # Cabeçalho destacado em tom azul escuro técnico
        pdf.set_fill_color(15, 23, 42)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.cell(95, 8, "  ITEM / SERVIÇO", border=0, ln=0, fill=True)
        pdf.cell(15, 8, "QTDE", border=0, ln=0, align="C", fill=True)
        pdf.cell(40, 8, "VALOR BASE   ", border=0, ln=0, align="R", fill=True)
        pdf.cell(40, 8, "À VISTA (PIX)   ", border=0, ln=1, align="R", fill=True)
        
        # Linhas de Itens com efeito zebra (fundo alternado sutil)
        pdf.set_font("Helvetica", "", 9.5)
        pdf.set_text_color(15, 23, 42)
        pdf.set_draw_color(241, 245, 249) # Bordas horizontais super leves
        
        total_base = 0.0
        total_avista = 0.0
        
        for idx, item in enumerate(st.session_state.lista_itens_orcamento):
            bg = (255, 255, 255) if idx % 2 == 0 else (248, 250, 252)
            pdf.set_fill_color(*bg)
            
            pdf.cell(95, 8, f"  {item['item']}", border="B", ln=0, fill=True)
            pdf.cell(15, 8, str(item['qtde']), border="B", ln=0, align="C", fill=True)
            pdf.cell(40, 8, f"R$ {item['base']:.2f}   ", border="B", ln=0, align="R", fill=True)
            pdf.cell(40, 8, f"R$ {item['avista']:.2f}   ", border="B", ln=1, align="R", fill=True)
            
            total_base += item['base']
            total_avista += item['avista']
            
        # Linha de Opções de Pagamento (Idêntica ao modelo integrado)
        pdf.set_fill_color(241, 245, 249)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(15, 23, 42)
        pdf.cell(95, 8, "  opções de pagamento", border="B", ln=0, fill=True)
        pdf.cell(15, 8, "", border="B", ln=0, fill=True)
        pdf.cell(40, 8, f"R$ {total_base:.2f}   ", border="B", ln=0, align="R", fill=True)
        pdf.cell(40, 8, f"R$ {total_avista:.2f}   ", border="B", ln=1, align="R", fill=True)
        
        # Simulação Dinâmica de Crédito Parcelado
        parcela_3x = total_base / 3
        pdf.set_fill_color(255, 255, 255)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(100, 116, 139)
        pdf.cell(95, 8, "  Cartão de Crédito (VISA / Mastercard / Hipercard em até 3x)", border="B", ln=0)
        pdf.cell(15, 8, "", border="B", ln=0)
        pdf.cell(40, 8, f"3x R$ {parcela_3x:.2f}   ", border="B", ln=0, align="R")
        pdf.cell(40, 8, "-", border="B", ln=1, align="R")
        pdf.ln(8)
        
        # Bloco de Termos de Garantia Limpo
        if termos_garantia:
            termos_limpos = termos_garantia.replace("•", "-")
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(15, 23, 42)
            pdf.cell(0, 5, "Garantia", ln=1)
            pdf.ln(1)
            pdf.set_font("Helvetica", "", 9)
            pdf.set_text_color(71, 85, 105)
            pdf.multi_cell(0, 4.5, termos_limpos)
            pdf.ln(10)
            
        # Sessão Inferior Estilizada de Assinaturas
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(148, 163, 184)
        pdf.cell(90, 5, "________________________________________", ln=0, align="C")
        pdf.cell(10, 5, "")
        pdf.cell(90, 5, "________________________________________", ln=1, align="C")
        pdf.set_text_color(71, 85, 105)
        pdf.cell(90, 5, "L.S.A Informática", ln=0, align="C")
        pdf.cell(10, 5, "")
        pdf.cell(90, 5, "Aceite do Cliente", ln=1, align="C")

        # Geração de saída estável
        nome_arquivo = f"Orcamento_Design_LSA_{cliente.replace(' ', '_')}.pdf"
        pdf.output(nome_arquivo)
        
        with open(nome_arquivo, "rb") as f:
            st.success("✨ PDF estruturado e compilado com matriz de blocos dinâmicos!")
            st.download_button("📥 Baixar Orçamento PDF Design", f, file_name=nome_arquivo, mime="application/pdf")