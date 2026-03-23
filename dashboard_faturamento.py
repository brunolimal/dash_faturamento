# -*- coding: utf-8 -*-
from flask import Flask, render_template_string, request
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import traceback
import locale
import json

# Conexão 100% Python que não precisa de driver da Microsoft no servidor (Render)
import pytds 

try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except:
    locale.setlocale(locale.LC_ALL, '') 

app = Flask(__name__)

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_k_m(valor):
    """Formata valores grandes para K (milhares) ou M (milhões) nos gráficos"""
    if pd.isna(valor) or valor == 0:
        return ""
    if valor >= 1000000:
        return f"R$ {valor/1000000:.1f}M"
    elif valor >= 1000:
        return f"R$ {valor/1000:.1f}K"
    return f"R$ {int(valor)}"

# Variáveis globais para o sistema de Cache (Memória Rápida)
_CACHE_DADOS = None
_CACHE_TEMPO = None

def carregar_dados():
    global _CACHE_DADOS, _CACHE_TEMPO
    
    # Cache de 10 minutos
    if _CACHE_DADOS is not None and _CACHE_TEMPO is not None:
        if datetime.now() - _CACHE_TEMPO < timedelta(minutes=10):
            print("🚀 Utilizando dados em cache...")
            return _CACHE_DADOS.copy()

    try:
        with pytds.connect(
            server='bi.srv.sisloc.com',
            user='dw_maisescoramentos',
            password='#45%maisWt',
            database='DW'
        ) as conn:
            with conn.cursor() as cursor:
                query = """
                SET NOCOUNT ON;
                EXEC DW_API '1A44894D6D3E39329B75F827426E2EA4',
                '
                SELECT 
                    nf.*, 
                    p_cliente.nm_pessoa AS nm_cliente
                FROM nf
                JOIN pessoa p_cliente ON nf.cd_pessoa = p_cliente.cd_pessoa
                '
                """
                cursor.execute(query)
                
                while cursor.description is None:
                    if not cursor.nextset():
                        break
                
                if cursor.description:
                    rows = cursor.fetchall()
                    colunas = [column[0] for column in cursor.description]
                    df = pd.DataFrame(rows, columns=colunas)
                else:
                    df = pd.DataFrame()

        if df.empty:
            return pd.DataFrame()

        df['vl_faturamento_bruto'] = pd.to_numeric(df['vl_faturamento_bruto'], errors='coerce').fillna(0)
        
        col_data = 'dt_emissao' if 'dt_emissao' in df.columns else df.filter(like='dt_').columns[0]
        df['dt_dashboard'] = pd.to_datetime(df[col_data], errors='coerce')
        df = df.dropna(subset=['dt_dashboard'])
        
        df['ano'] = df['dt_dashboard'].dt.year
        df['mes_num'] = df['dt_dashboard'].dt.month
        
        if 'fl_origem' in df.columns:
            df['origem_upper'] = df['fl_origem'].astype(str).str.strip().str.upper()
            tags_venda = ['VD', 'DV', 'IL', 'VENDA DE LOCAÇÃO']
            tags_locacao = ['FL', 'SL']
            df['is_venda'] = df['origem_upper'].isin(tags_venda)
            df['is_locacao'] = df['origem_upper'].isin(tags_locacao)
            
            # Criar coluna unificada para legenda
            df['Tipo'] = 'Outros'
            df.loc[df['is_venda'], 'Tipo'] = 'Venda'
            df.loc[df['is_locacao'], 'Tipo'] = 'Locação'
        else:
            df['is_venda'] = False
            df['is_locacao'] = False
            df['Tipo'] = 'Indefinido'

        _CACHE_DADOS = df.copy()
        _CACHE_TEMPO = datetime.now()

        return df

    except Exception as e:
        print("\n❌ Erro crítico:")
        traceback.print_exc()
        raise e

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta http-equiv="refresh" content="600">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mais Escoramentos | BI Faturamento</title>
    
    <!-- Google Fonts & Tailwind CSS -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: { sans: ['Inter', 'sans-serif'], },
                    colors: {
                        brand: {
                            dark: '#0f172a',
                            blue: '#0ea5e9',
                            green: '#10b981',
                            indigo: '#6366f1',
                            orange: '#f59e0b'
                        }
                    }
                }
            }
        }
    </script>
    
    <style>
        body { background-color: #f8fafc; }
        .glass-panel { background: white; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03); border-radius: 1rem; }
        .chart-container { width: 100%; height: 100%; min-height: 350px; }
        
        #loader {
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(255,255,255,0.85); backdrop-filter: blur(4px); z-index: 9999; 
            justify-content: center; align-items: center; flex-direction: column;
        }
        .spinner { border: 4px solid #f1f5f9; border-top: 4px solid #0ea5e9; border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
        
        .pulse-dot { width: 8px; height: 8px; background-color: #10b981; border-radius: 50%; animation: pulse 2s infinite; }
        @keyframes pulse { 0% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); } 70% { box-shadow: 0 0 0 6px rgba(16, 185, 129, 0); } 100% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); } }
    </style>
</head>
<body class="text-slate-800 antialiased p-4 md:p-8">

    <!-- Loader -->
    <div id="loader">
        <div class="spinner"></div>
        <h3 class="mt-4 font-semibold text-slate-700">Atualizando BI...</h3>
    </div>

    <div class="max-w-7xl mx-auto space-y-6">
        
        <!-- Header & Filters -->
        <div class="flex flex-col lg:flex-row justify-between items-start lg:items-center bg-brand-dark rounded-2xl p-6 lg:p-8 shadow-lg text-white gap-6">
            <div>
                <div class="flex items-center gap-3">
                    <div class="bg-blue-500/20 p-2 rounded-lg">
                        <i class="fa-solid fa-chart-pie text-brand-blue text-2xl"></i>
                    </div>
                    <h1 class="text-2xl md:text-3xl font-bold tracking-tight">Faturamento Escoramentos</h1>
                </div>
                <p class="text-slate-400 mt-2 text-sm font-medium">Análise e Acompanhamento de Resultados</p>
            </div>
            
            <form method="GET" class="flex flex-wrap items-center gap-4 bg-white/10 p-3 rounded-xl border border-white/10 backdrop-blur-sm" id="form-filtros">
                <div class="flex items-center gap-2 text-sm font-semibold text-slate-300 uppercase tracking-wider mr-2">
                    <i class="fa-solid fa-filter"></i> Filtros
                </div>
                
                <select name="ano" onchange="mostrarLoaderEEnviar()" class="bg-white text-brand-dark px-4 py-2 rounded-lg font-semibold text-sm outline-none cursor-pointer focus:ring-2 focus:ring-brand-blue">
                    {% for a in anos %}
                        <option value="{{ a }}" {% if a|string == ano_sel %}selected{% endif %}>Ano: {{ a }}</option>
                    {% endfor %}
                </select>
                
                <select name="mes" onchange="mostrarLoaderEEnviar()" class="bg-white text-brand-dark px-4 py-2 rounded-lg font-semibold text-sm outline-none cursor-pointer focus:ring-2 focus:ring-brand-blue">
                    {% for v, n in meses %}
                        <option value="{{ v }}" {% if v|string == mes_sel %}selected{% endif %}>Mês: {{ n }}</option>
                    {% endfor %}
                </select>
            </form>
        </div>

        <!-- KPI Cards -->
        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
            <!-- Total -->
            <div class="glass-panel p-6 relative overflow-hidden group">
                <div class="absolute -right-4 -top-4 opacity-5 group-hover:scale-110 transition-transform duration-300">
                    <i class="fa-solid fa-wallet text-9xl"></i>
                </div>
                <div class="flex justify-between items-start mb-4">
                    <div class="bg-blue-100 text-brand-blue p-3 rounded-lg"><i class="fa-solid fa-dollar-sign text-xl"></i></div>
                    <span class="text-xs font-bold px-2 py-1 bg-slate-100 text-slate-500 rounded-full">Bruto</span>
                </div>
                <p class="text-sm font-bold text-slate-400 uppercase tracking-wider">Faturamento Total</p>
                <h3 class="text-2xl md:text-3xl font-black text-slate-800 mt-1">{{ total_faturamento }}</h3>
            </div>

            <!-- Locação -->
            <div class="glass-panel p-6 relative overflow-hidden group">
                <div class="absolute -right-4 -top-4 opacity-5 group-hover:scale-110 transition-transform duration-300">
                    <i class="fa-solid fa-truck-ramp-box text-9xl"></i>
                </div>
                <div class="flex justify-between items-start mb-4">
                    <div class="bg-green-100 text-brand-green p-3 rounded-lg"><i class="fa-solid fa-retweet text-xl"></i></div>
                    <span class="text-xs font-bold px-2 py-1 bg-green-50 text-brand-green rounded-full">{{ perc_locacao }}%</span>
                </div>
                <p class="text-sm font-bold text-slate-400 uppercase tracking-wider">Locação</p>
                <h3 class="text-2xl md:text-3xl font-black text-slate-800 mt-1">{{ fat_locacao }}</h3>
            </div>

            <!-- Vendas -->
            <div class="glass-panel p-6 relative overflow-hidden group">
                <div class="absolute -right-4 -top-4 opacity-5 group-hover:scale-110 transition-transform duration-300">
                    <i class="fa-solid fa-tags text-9xl"></i>
                </div>
                <div class="flex justify-between items-start mb-4">
                    <div class="bg-indigo-100 text-brand-indigo p-3 rounded-lg"><i class="fa-solid fa-cart-shopping text-xl"></i></div>
                    <span class="text-xs font-bold px-2 py-1 bg-indigo-50 text-brand-indigo rounded-full">{{ perc_vendas }}%</span>
                </div>
                <p class="text-sm font-bold text-slate-400 uppercase tracking-wider">Vendas</p>
                <h3 class="text-2xl md:text-3xl font-black text-slate-800 mt-1">{{ fat_vendas }}</h3>
            </div>

            <!-- Ticket Médio -->
            <div class="glass-panel p-6 relative overflow-hidden group">
                <div class="absolute -right-4 -top-4 opacity-5 group-hover:scale-110 transition-transform duration-300">
                    <i class="fa-solid fa-receipt text-9xl"></i>
                </div>
                <div class="flex justify-between items-start mb-4">
                    <div class="bg-orange-100 text-brand-orange p-3 rounded-lg"><i class="fa-solid fa-calculator text-xl"></i></div>
                    <span class="text-xs font-bold px-2 py-1 bg-slate-100 text-slate-500 rounded-full">{{ qtd_nfs }} NFs</span>
                </div>
                <p class="text-sm font-bold text-slate-400 uppercase tracking-wider">Ticket Médio</p>
                <h3 class="text-2xl md:text-3xl font-black text-slate-800 mt-1">{{ media }}</h3>
            </div>
        </div>

        <!-- Charts Row 1 -->
        <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
            <!-- Gráfico Mensal -->
            <div class="glass-panel p-6 lg:col-span-2 flex flex-col">
                <div class="flex justify-between items-center mb-4">
                    <h2 class="text-lg font-bold text-slate-800">Evolução Mensal</h2>
                    <span class="text-sm font-medium text-slate-500">Ano: {{ ano_sel if ano_sel != 'all' else 'Todos' }}</span>
                </div>
                <div class="flex-grow w-full">
                    <div id="g_mensal" class="chart-container"></div>
                </div>
            </div>

            <!-- Gráfico de Composição (Donut) -->
            <div class="glass-panel p-6 flex flex-col">
                <h2 class="text-lg font-bold text-slate-800 mb-4">Composição de Receita</h2>
                <div class="flex-grow w-full flex items-center justify-center">
                    <div id="g_composicao" class="chart-container"></div>
                </div>
            </div>
        </div>

        <!-- Charts Row 2 -->
        <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <!-- Top Clientes -->
            <div class="glass-panel p-6 flex flex-col">
                <h2 class="text-lg font-bold text-slate-800 mb-4">Top 10 Clientes</h2>
                <div class="flex-grow w-full">
                    <div id="g_top" class="chart-container" style="min-height: 400px;"></div>
                </div>
            </div>

            <!-- Gráfico Anual -->
            <div class="glass-panel p-6 flex flex-col">
                <div class="flex justify-between items-center mb-4">
                    <h2 class="text-lg font-bold text-slate-800">Histórico Anual</h2>
                    <span class="text-sm font-medium text-slate-500">Geral</span>
                </div>
                <div class="flex-grow w-full">
                    <div id="g_anual" class="chart-container" style="min-height: 400px;"></div>
                </div>
            </div>
        </div>

        <!-- Footer -->
        <div class="flex items-center justify-center gap-2 text-xs font-semibold text-slate-400 py-6">
            <div class="pulse-dot"></div>
            <span>Conectado ao DW • Última extração: {{ data_extracao }} • Atualização automática (10 min)</span>
        </div>

    </div>

    <script>
        function mostrarLoaderEEnviar() {
            document.getElementById('loader').style.display = 'flex';
            document.getElementById('form-filtros').submit();
        }

        // Configuração Padrão do Plotly
        var config = { responsive: true, displayModeBar: false };
        
        // Dados injetados do Python
        var g_mensal_data = {{ fig_mensal_json | safe }};
        var g_composicao_data = {{ fig_composicao_json | safe }};
        var g_anual_data = {{ fig_anual_json | safe }};
        var g_top_data = {{ fig_top_json | safe }};
        
        // Renderização dos Gráficos
        Plotly.newPlot('g_mensal', g_mensal_data.data, g_mensal_data.layout, config);
        Plotly.newPlot('g_composicao', g_composicao_data.data, g_composicao_data.layout, config);
        Plotly.newPlot('g_anual', g_anual_data.data, g_anual_data.layout, config);
        Plotly.newPlot('g_top', g_top_data.data, g_top_data.layout, config);
    </script>
</body>
</html>
"""

@app.route("/", methods=["GET"])
def dashboard():
    try:
        df = carregar_dados()
        if df.empty:
            return "<h2>⚠️ O banco não retornou dados.</h2>", 200
    except Exception as e:
        return f"<h2>⚠️ Erro ao conectar com o banco:</h2><p>{str(e)}</p>", 500

    anos = sorted(df["ano"].unique().tolist(), reverse=True)
    ano_padrao = str(anos[0]) if anos else "all"
    ano_sel = request.args.get("ano", ano_padrao)
    mes_sel = request.args.get("mes", "all")

    # -------------------------------------------------------------
    # 1. CÁLCULOS DOS CARDS (Usa o filtro completo)
    # -------------------------------------------------------------
    df_f = df.copy()
    if ano_sel != "all": df_f = df_f[df_f["ano"] == int(ano_sel)]
    if mes_sel != "all": df_f = df_f[df_f["mes_num"] == int(mes_sel)]

    total_faturamento = df_f["vl_faturamento_bruto"].sum()
    qtd_nfs = len(df_f)
    media = total_faturamento / qtd_nfs if qtd_nfs > 0 else 0
    
    faturamento_vendas = df_f[df_f['is_venda']]["vl_faturamento_bruto"].sum()
    faturamento_locacao = df_f[df_f['is_locacao']]["vl_faturamento_bruto"].sum()
    
    # Percentuais
    perc_vendas = round((faturamento_vendas / total_faturamento * 100), 1) if total_faturamento > 0 else 0
    perc_locacao = round((faturamento_locacao / total_faturamento * 100), 1) if total_faturamento > 0 else 0

    # Layout Base Moderno para Plotly
    layout_moderno = dict(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Inter, sans-serif", color="#64748b", size=12),
        margin=dict(l=0, r=0, t=10, b=20),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )

    meses_map = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
    ordem_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    cores_tipo = {'Locação': '#10b981', 'Venda': '#6366f1', 'Outros': '#cbd5e1'}

    # -------------------------------------------------------------
    # 2. GRÁFICO: Evolução Mensal (Barras Empilhadas)
    # -------------------------------------------------------------
    df_mensal = df.copy()
    if ano_sel != "all": df_mensal = df_mensal[df_mensal["ano"] == int(ano_sel)]
    
    # Agrupar por Mês e Tipo
    df_mensal_grp = df_mensal[df_mensal['Tipo'] != 'Outros'].groupby(['mes_num', 'Tipo'])["vl_faturamento_bruto"].sum().reset_index()
    df_mensal_grp['mes_nome'] = df_mensal_grp['mes_num'].map(meses_map)
    df_mensal_grp['rotulo'] = df_mensal_grp['vl_faturamento_bruto'].apply(formatar_k_m)

    fig_mensal = px.bar(
        df_mensal_grp, x="mes_nome", y="vl_faturamento_bruto", color="Tipo", text="rotulo",
        color_discrete_map=cores_tipo, barmode="stack"
    )
    fig_mensal.update_traces(textposition="inside", insidetextanchor="middle", textfont_color="white")
    fig_mensal.update_layout(**layout_moderno)
    fig_mensal.update_yaxes(visible=False, showticklabels=False)
    fig_mensal.update_xaxes(title="", categoryorder='array', categoryarray=ordem_meses)

    # -------------------------------------------------------------
    # 3. GRÁFICO: Composição (Rosca/Donut)
    # -------------------------------------------------------------
    df_comp = df_f[df_f['Tipo'] != 'Outros'].groupby('Tipo')["vl_faturamento_bruto"].sum().reset_index()
    
    fig_composicao = px.pie(
        df_comp, values='vl_faturamento_bruto', names='Tipo', hole=0.6,
        color='Tipo', color_discrete_map=cores_tipo
    )
    fig_composicao.update_traces(textinfo='percent', hoverinfo='label+value', textfont_size=14, textfont_color="white", marker=dict(line=dict(color='#ffffff', width=2)))
    fig_composicao.update_layout(
        **layout_moderno,
        margin=dict(l=0, r=0, t=20, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5)
    )
    # Colocar o Total no centro do Donut
    fig_composicao.add_annotation(
        text=formatar_moeda(df_comp["vl_faturamento_bruto"].sum()).replace("R$ ", ""), 
        x=0.5, y=0.5, font_size=20, font_family="Inter", font_weight="bold", showarrow=False
    )

    # -------------------------------------------------------------
    # 4. GRÁFICO: Histórico Anual (Barras Agrupadas)
    # -------------------------------------------------------------
    df_anual = df.copy()
    if mes_sel != "all": df_anual = df_anual[df_anual["mes_num"] == int(mes_sel)]
    
    df_anual_grp = df_anual[df_anual['Tipo'] != 'Outros'].groupby(['ano', 'Tipo'])["vl_faturamento_bruto"].sum().reset_index()
    df_anual_grp['rotulo'] = df_anual_grp['vl_faturamento_bruto'].apply(formatar_k_m)
    df_anual_grp['ano'] = df_anual_grp['ano'].astype(str)

    fig_anual = px.bar(
        df_anual_grp, x="ano", y="vl_faturamento_bruto", color="Tipo", text="rotulo",
        color_discrete_map=cores_tipo, barmode="group"
    )
    fig_anual.update_traces(textposition="outside")
    fig_anual.update_layout(**layout_moderno, margin=dict(l=0, r=0, t=30, b=20))
    fig_anual.update_yaxes(visible=False, showticklabels=False)
    fig_anual.update_xaxes(title="", type='category', categoryorder='category ascending')

    # -------------------------------------------------------------
    # 5. GRÁFICO: Top 10 Clientes
    # -------------------------------------------------------------
    df_top = df_f.groupby("nm_cliente")["vl_faturamento_bruto"].sum().nlargest(10).reset_index()
    df_top['rotulo'] = df_top['vl_faturamento_bruto'].apply(formatar_k_m)
    
    fig_top = px.bar(df_top, x="vl_faturamento_bruto", y="nm_cliente", orientation='h', text="rotulo")
    fig_top.update_traces(marker_color="#0ea5e9", textposition="outside") # Azul padrão para o Top
    fig_top.update_layout(yaxis={'categoryorder':'total ascending'}, **layout_moderno, margin=dict(l=10, r=40, t=10, b=10))
    fig_top.update_xaxes(visible=False, showticklabels=False)
    fig_top.update_yaxes(title="", tickfont=dict(size=10))

    # Lista de meses para o Filtro HTML
    meses = [("all", "Todos os Meses"), (1, "Janeiro"), (2, "Fevereiro"), (3, "Março"), (4, "Abril"), (5, "Maio"), (6, "Junho"), 
             (7, "Julho"), (8, "Agosto"), (9, "Setembro"), (10, "Outubro"), (11, "Novembro"), (12, "Dezembro")]

    return render_template_string(HTML_TEMPLATE,
        data_extracao=datetime.now().strftime('%d/%m/%Y às %H:%M'),
        anos=anos,
        ano_sel=ano_sel,
        meses=meses,
        mes_sel=mes_sel,
        total_faturamento=formatar_moeda(total_faturamento),
        fat_vendas=formatar_moeda(faturamento_vendas),
        fat_locacao=formatar_moeda(faturamento_locacao),
        perc_vendas=perc_vendas,
        perc_locacao=perc_locacao,
        media=formatar_moeda(media),
        qtd_nfs=qtd_nfs,
        fig_mensal_json=fig_mensal.to_json(),
        fig_composicao_json=fig_composicao.to_json(),
        fig_anual_json=fig_anual.to_json(),
        fig_top_json=fig_top.to_json()
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
