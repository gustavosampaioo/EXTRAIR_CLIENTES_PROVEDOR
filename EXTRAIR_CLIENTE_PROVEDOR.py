import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import re

def processar_coordenada(coord_str):
    """Processa a string de coordenada do formato do Excel"""
    try:
        # Remove espaços extras e divide a coordenada
        coord = str(coord_str).strip()
        # Substitui vírgula por ponto se necessário
        coord = coord.replace(',', '.')
        # Divide por espaço
        partes = coord.split()
        if len(partes) >= 2:
            # Assume formato: latitude longitude
            lat = partes[0]
            lon = partes[1]
            return f"{lon},{lat},0"  # Formato KML: longitude,latitude,altitude
    except:
        return None
    return None

def extrair_nome_equipamento(ipv4):
    """Extrai o texto antes do ' - ' da coluna Ipv4"""
    try:
        if pd.notna(ipv4) and ' - ' in str(ipv4):
            return str(ipv4).split(' - ')[0].strip()
    except:
        pass
    return str(ipv4) if pd.notna(ipv4) else ''

def extrair_interface_conexao(interface):
    """Extrai apenas a parte final 1/16/15 da coluna Interface Conexão"""
    try:
        if pd.notna(interface):
            texto = str(interface)
            # Procura por padrão de números separados por /
            padrao = r'(\d+/\d+/\d+)'
            match = re.search(padrao, texto)
            if match:
                return match.group(1)
    except:
        pass
    return str(interface) if pd.notna(interface) else ''

def gerar_kml(df):
    """Gera o conteúdo KML a partir do DataFrame"""
    
    # Cabeçalho do KML
    kml_header = '''<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Pontos de Equipamentos</name>
'''
    
    # Estilo padrão para os placemarks
    kml_style = '''
    <Style id="style_ponto">
      <IconStyle>
        <scale>1.2</scale>
        <Icon>
          <href>http://maps.google.com/mapfiles/kml/pushpin/red-pushpin.png</href>
        </Icon>
      </IconStyle>
    </Style>
'''
    
    kml_body = kml_header + kml_style
    placemarks_count = 0
    
    # Processar cada linha do DataFrame
    for idx, row in df.iterrows():
        # Processar coordenada da coluna G (índice 6)
        coordenada = processar_coordenada(row.iloc[6] if len(row) > 6 else None)
        
        if coordenada:
            placemarks_count += 1
            
            # Extrair informações das colunas (A até F)
            cliente = str(row.iloc[0]) if len(row) > 0 and pd.notna(row.iloc[0]) else ''
            estado = str(row.iloc[1]) if len(row) > 1 and pd.notna(row.iloc[1]) else ''
            cidade = str(row.iloc[2]) if len(row) > 2 and pd.notna(row.iloc[2]) else ''
            bairro = str(row.iloc[3]) if len(row) > 3 and pd.notna(row.iloc[3]) else ''
            endereco = str(row.iloc[4]) if len(row) > 4 and pd.notna(row.iloc[4]) else ''
            
            # Login (coluna F - índice 5)
            login = str(row.iloc[5]) if len(row) > 5 and pd.notna(row.iloc[5]) else ''
            
            # Interface Conexão (coluna H - índice 7)
            interface_raw = row.iloc[7] if len(row) > 7 else ''
            interface_conexao = extrair_interface_conexao(interface_raw)
            
            # Processar Ipv4 (coluna I - índice 8)
            ipv4_raw = row.iloc[8] if len(row) > 8 else ''
            ipv4_nome = extrair_nome_equipamento(ipv4_raw)
            
            # RX Signal (coluna J - índice 9)
            rx_signal = str(row.iloc[9]) if len(row) > 9 and pd.notna(row.iloc[9]) else ''
            
            # Nome do placemark (Ipv4 + Interface Conexão)
            placemark_name = f"{ipv4_nome} - {interface_conexao}"
            
            # Descrição do placemark com todas as informações
            description = f"""
            <![CDATA[
            <table style="font-family: Arial, sans-serif; border-collapse: collapse; width: 100%;">
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Cliente:</b></td><td style="padding: 5px;">{cliente}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Estado:</b></td><td style="padding: 5px;">{estado}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Cidade:</b></td><td style="padding: 5px;">{cidade}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Bairro:</b></td><td style="padding: 5px;">{bairro}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Endereço:</b></td><td style="padding: 5px;">{endereco}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Login:</b></td><td style="padding: 5px;">{login}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Interface Conexão:</b></td><td style="padding: 5px;">{interface_conexao}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>Ipv4:</b></td><td style="padding: 5px;">{ipv4_nome}</td></tr>
                <tr><td style="padding: 5px; background-color: #f2f2f2;"><b>RX Signal:</b></td><td style="padding: 5px;">{rx_signal}</td></tr>
            </table>
            ]]>
            """
            
            # Adicionar placemark ao KML
            placemark = f'''
    <Placemark>
      <name>{placemark_name}</name>
      <description>{description}</description>
      <styleUrl>#style_ponto</styleUrl>
      <Point>
        <coordinates>{coordenada}</coordinates>
      </Point>
    </Placemark>
'''
            kml_body += placemark
    
    # Rodapé do KML
    kml_footer = '''
  </Document>
</kml>'''
    
    kml_body += kml_footer
    return kml_body, placemarks_count

def main():
    st.set_page_config(
        page_title="Gerador de KML",
        page_icon="📍",
        layout="centered"
    )
    
    st.title("📍 Gerador de KML a partir de Planilha Excel")
    st.markdown("---")
    
    st.markdown("""
    ### Instruções:
    1. Faça upload de um arquivo Excel (.xlsx, .xls)
    2. O arquivo deve conter as colunas na seguinte ordem:
        - **Coluna A**: Cliente
        - **Coluna B**: Estado
        - **Coluna C**: Cidade
        - **Coluna D**: Bairro
        - **Coluna E**: Endereço
        - **Coluna F**: Login
        - **Coluna G**: Coordenada (formato: -5.129502 -42.781442)
        - **Coluna H**: Interface Conexão (ex: [GPON] SLOT 16 - PON 15 - 1/16/15)
        - **Coluna I**: Ipv4 (ex: PS - 10.252.0.2)
        - **Coluna J**: RX Signal (ex: -25.5 dBm)
    3. Clique em "Gerar KML" para fazer o download
    """)
    
    st.markdown("---")
    
    # Upload do arquivo
    uploaded_file = st.file_uploader(
        "Escolha um arquivo Excel",
        type=['xlsx', 'xls'],
        help="Selecione o arquivo Excel com os dados dos pontos"
    )
    
    if uploaded_file is not None:
        try:
            # Ler o arquivo Excel sem cabeçalho
            df = pd.read_excel(uploaded_file, header=None)
            
            # Mostrar preview dos dados
            st.subheader("📊 Preview dos dados (primeiras 5 linhas)")
            
            # Criar um DataFrame com cabeçalhos para melhor visualização
            preview_df = df.head(5).copy()
            preview_df.columns = ['Cliente', 'Estado', 'Cidade', 'Bairro', 'Endereço', 
                                 'Login', 'Coordenada', 'Interface Conexão', 'Ipv4', 'RX Signal']
            st.dataframe(preview_df)
            
            # Mostrar número de linhas
            st.info(f"Total de linhas no arquivo: {len(df)}")
            
            # Verificar se há coordenadas válidas na coluna G (índice 6)
            coordenadas_validas = 0
            for idx, row in df.iterrows():
                if len(row) > 6 and processar_coordenada(row.iloc[6]):
                    coordenadas_validas += 1
            
            if coordenadas_validas == 0:
                st.warning("⚠️ Nenhuma coordenada válida encontrada na coluna G. Verifique o formato das coordenadas.")
            else:
                st.success(f"✅ {coordenadas_validas} coordenadas válidas encontradas")
            
            # Botão para gerar KML
            if st.button("🔄 Gerar KML", type="primary", use_container_width=True):
                with st.spinner("Gerando arquivo KML..."):
                    # Gerar o KML
                    kml_content, placemarks_count = gerar_kml(df)
                    
                    # Criar arquivo para download
                    kml_bytes = kml_content.encode('utf-8')
                    
                    # Nome do arquivo
                    filename = "pontos_equipamentos.kml"
                    
                    # Mensagem de sucesso
                    st.success(f"✅ KML gerado com sucesso! {placemarks_count} placemarks criados.")
                    
                    # Botão de download
                    st.download_button(
                        label="📥 Download KML",
                        data=kml_bytes,
                        file_name=filename,
                        mime="application/vnd.google-earth.kml+xml",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Erro ao processar o arquivo: {str(e)}")
            st.exception(e)
    
    st.markdown("---")
    st.markdown("### 📝 Exemplo do formato das colunas (ordem correta):")
    
    # Criar um DataFrame de exemplo para mostrar a ordem das colunas
    exemplo_data = {
        'Coluna': ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'],
        'Campo': ['Cliente', 'Estado', 'Cidade', 'Bairro', 'Endereço', 'Login', 
                 'Coordenada', 'Interface Conexão', 'Ipv4', 'RX Signal'],
        'Exemplo': ['João Silva', 'PI', 'Teresina', 'Centro', 'Rua Principal, 123', 
                   'joao.silva', '-5.129502 -42.781442', '[GPON] SLOT 16 - PON 15 - 1/16/15', 
                   'PS - 10.252.0.2', '-25.5 dBm']
    }
    exemplo_df = pd.DataFrame(exemplo_data)
    st.dataframe(exemplo_df, use_container_width=True)
    
    st.markdown("""
    **Observações:**
    - O nome do placemark será formado por: `Ipv4 - Interface Conexão` (ex: PS - 1/16/15)
    - A coordenada deve estar no formato: `latitude longitude` (ex: -5.129502 -42.781442)
    - Todas as informações (Cliente, Estado, Cidade, Bairro, Endereço, Login, Interface Conexão, Ipv4 e RX Signal) serão incluídas na descrição do ponto no KML
    """)

if __name__ == "__main__":
    main()
