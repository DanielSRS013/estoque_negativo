import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from pathlib import Path

lojas = ['LNY BARRA',
         'LNY BH',
         'LNY BRASILIA',
         'LNY BUZIOS',
         'LNY CAMPINAS',
         'LNY CURITIBA',
         'LNY FASHION MALL',
         'LNY IPANEMA',
         'LNY GARCIA',
         'LNY GOIANIA',
         'LNY HIGIENOPOLIS',
         'LNY JARDINS',
         'LNY NITEROI',
         'LNY OFF CATARINA',
         'LNY OFF LEBLON',
         'LNY PORTO ALEGRE',
         'LNY RECIFE',
         'LNY RIBEIRÃO PRETO',
         'LNY RIO SUL',
         'LNY SALVADOR',
         'LNY SALVADOR SHOPPING',
         'LNY SAVASSI',
         'LNY LEBLON',
         'LNY TRANCOSO',
         'LNY VILLAGE MALL',
         'LNY VITORIA'
]


if 'produtos' not in st.session_state:
    st.session_state.produtos = []

EMAIL_USER = st.secrets["EMAIL_USER"]
EMAIL_PASS = st.secrets["EMAIL_PASS"]

def send_email(df_excel, loja):
    email = EmailMessage()
    email['Subject'] = f'FORMULÁRIO ESTOQUE NEGATIVO {loja}'
    email['From'] = EMAIL_USER
    email['To'] = 'aprendiz.auditoria@lennyniemeyer.com'
    email['Cc'] = 'alice.costa@lennyniemeyer.com; girlene.silva@lennyniemeyer.com'

    email.set_content("""Olá,
Segue em anexo o formulário de estoque em formato Excel.
Atenciosamente,
Sistema Automático
""")
    
    caminho = Path(df_excel)

    with open(caminho, 'rb') as f:
        email.add_attachment(
            f.read(),
            maintype = 'application',
            subtype = 'vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            filename = caminho.name)

    with smtplib.SMTP("smtp.gmail.com", 587) as smtp:
        smtp.starttls()
        smtp.login(EMAIL_USER, EMAIL_PASS)
        smtp.send_message(email)


with st.form('my-form', clear_on_submit=True):

    st.write('Formulário')
    loja = st.selectbox(label = 'Digite a Loja',index = None, placeholder='', options = lojas)
    name = st.text_input(label ='Digite o Nome do produto', key = 'name')
    reference = st.text_input('Digite a Referência', key = 'reference')
    color = st.text_input('Digite a Cor (Código)', key = 'color')
    size = st.text_input('Digite o Tamanho', key = 'size')
    fisic_storage = st.text_input('Digite o valor identificado na contagem', key = 'fisic_storage')

    ref_answer = st.selectbox('Referência correta?', ['Sim'], index = None, placeholder='Sim ou digite a referência que foi movimentada', accept_new_options=True)
    size_answer = st.selectbox('Tamanho Correto?', ['Sim'], index = None, placeholder='Sim ou digite o tamanho que foi movimentado', accept_new_options=True)    
    print_answer = st.selectbox('Estampa Correta?', ['Sim'], index = None, placeholder ='Sim ou digite a estampa que foi movimentada')
    invoice_answer = st.selectbox('Pendência de Nota Fiscal?', ['Não', 'Sim'])
    delivery_answer = st.selectbox('Consta em Delivery?', ['Não', 'Sim'])
   
    col1, col2, col3 = st.columns(3)
    add_product = col1.form_submit_button('Adicionar Produto')
    send_form = col3.form_submit_button('Enviar Formulário')

data = [loja, name, reference, color, size, fisic_storage, ref_answer , size_answer, print_answer, invoice_answer, delivery_answer]

if add_product:
    if not all(d and str(d).strip() for d in data):
        st.error('Preencher todos os campos')
    else:
        st.session_state.produtos.append({'Loja': loja,
            'Nome': name, 'Referência':reference, 'Cor':color, 'Tamanho':size, 'Estoque físico': fisic_storage, 'Referência correta': ref_answer, 'Tamanho Correto': size_answer, 'Estampa Correta': print_answer, 'Pendência de Nota Fiscal': invoice_answer, 'Consta em Delivery':delivery_answer
        })
if st.session_state.produtos:
    produtos = st.session_state.produtos  
    ref_selecionada = st.selectbox(
        'Selecione o produto para alterar',
        options = [p['Referência'] for p in produtos],
        key = 'ref_selecionada')
    if ref_selecionada:
        produto = next(p for p in produtos if p['Referência'] == ref_selecionada)
        with st.form('editar_produto'):
            loja = st.text_input('Loja', value= produto['Loja'])
            name = st.text_input('Nome', value = produto['Nome'])
            reference = st.text_input('Referência', value = produto['Referência'], disabled=True)
            color = st.text_input('Cor', value = produto['Cor'])
            size = st.text_input('Tamanho', value = produto['Tamanho'])
            fisic_storage = st.text_input('Estoque físico', value = produto['Estoque físico'])               
            ref_answer = st.selectbox('Referência correta', ['Sim', 'Não'], 
            index = 0 if produto['Referência correta'] == 'Sim' else 1)
            size_answer = st.selectbox(
                'Tamanho Correto', 
                ['Sim', 'Não'],
                index = 0 if produto['Tamanho Correto'] == 'Sim' else 1)               
            print_answer = st.selectbox('Estampa Correta',['Sim', 'Não'], index = 0 if produto['Estampa Correta']=='Sim' else 1)
            invoice_answer = st.selectbox('Pendência de Nota Fiscal',['Sim', 'Não'], index = 0 if produto['Pendência de Nota Fiscal'] == 'Sim' else 1)
            delivery_answer = st.selectbox('Consta em Delivery',['Sim', 'Não'], index = 0 if produto['Consta em Delivery'] == 'Sim' else 1)
            save = st.form_submit_button('Salvar alterações')
            
        if save:
            produto.update({'Loja': loja,
            'Nome': name, 'Referência':reference, 'Cor':color, 'Tamanho':size, 'Estoque físico': fisic_storage, 'Referência correta': ref_answer, 'Tamanho Correto': size_answer, 'Estampa Correta': print_answer, 'Pendência de Nota Fiscal': invoice_answer, 'Consta em Delivery':delivery_answer})
            st.success('Produto atualizado com sucesso!')

    if send_form:

        if not produtos:
            st.warning('Nenhum produto adicionado')
            st.stop()

        else:         
            df_produtos = pd.DataFrame(produtos)
            arquivo = 'formulario-estoque.xlsx'
            df_produtos_loja = df_produtos['Loja'].unique()
            df_produtos.to_excel(arquivo, index=False)
            send_email(df_excel = arquivo, loja = df_produtos_loja[0])
            st.success('E-mail enviado com sucesso!')