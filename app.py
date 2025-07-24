import streamlit as st
import pandas as pd
from datetime import datetime
import random

st.set_page_config(page_title="Protótipo AD - Cadastro por Planilha", page_icon="📄")

st.title("📄 Protótipo de Cadastro de Usuários no AD a partir de Planilha")

uploaded_file = st.file_uploader("📎 Envie a planilha (.xlsx) com os dados do RH", type="xlsx")

def gerar_login(nome_completo):
    partes = nome_completo.strip().split()
    if len(partes) < 3:
        return None, None, None, None
    sigla = partes[0][0] + partes[-2][0] + partes[-1][0]
    login = (partes[0][0] + partes[-2][0] + partes[-1]).lower()
    first_name = f"[{sigla.upper()}] {partes[0]}"
    last_name = " ".join(partes[1:])
    full_name = f"{first_name} {last_name}"
    return sigla.upper(), login, first_name, last_name, full_name

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.write("📋 Dados lidos da planilha:")
    st.dataframe(df)

    st.subheader("👨‍💻 Usuários processados:")

    for idx, row in df.iterrows():
        nome = row['Nome Completo']
        cargo = row['Cargo']
        unidade = row['Unidade']
        gestor = row['Gestor']
        setor = row['Setor']
        usuario_espelho = row['Usuario Espelho']

        sigla, login, first_name, last_name, full_name = gerar_login(nome)

        if not login:
            st.warning(f"❌ Nome inválido ou incompleto: {nome}")
            continue

        email = f"{login}@almeidaadvogados.com.br"
        senha = "Almeida@2024@"
        protocolo = random.randint(10000, 99999)
        data = datetime.now().strftime("%d/%m/%Y %H:%M")

        with st.expander(f"✅ Usuário {nome} (Protocolo #{protocolo})"):
            st.markdown(f"""- **Nome Completo:** {nome}  
- **Login (sAMAccountName):** `{login}`  
- **Sigla:** `{sigla}`  
- **First Name:** `{first_name}`  
- **Last Name:** `{last_name}`  
- **Full Name:** `{full_name}`  
- **Cargo (Description):** {cargo}  
- **Unidade (Office):** {unidade}  
- **Gestor:** {gestor}  
- **Setor:** {setor}  
- **Usuário Espelho:** {usuario_espelho.upper()}  
- **Email:** `{email}`  
- **Senha inicial:** `{senha}`  
- **Data/Hora:** {data}  
- 🔐 **Usuário deve trocar senha no primeiro login**""")
