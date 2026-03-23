# -*- coding: utf-8 -*-
from flask import Flask, render_template_string, request
import pandas as pd
import plotly.express as px
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
    
    # Se os dados já estiverem na memória e tiverem menos de 10 minutos, usamos os dados guardados!
    if _CACHE_DADOS is not None and _CACHE_TEMPO is not None:
        if datetime.now() - _CACHE_TEMPO < timedelta(minutes=10):
            print("🚀 A usar dados da memória (Cache)!")
            return _CACHE_DADOS.copy()

    try:
        # Usando a biblioteca pytds.
        with pytds.connect(
            server='bi.srv.sisloc.com',
            user='dw_maisescoramentos',
            password='#45%maisWt',
            database='DW'
        ) as conn:
            with conn.cursor() as cursor:
                # 1ª Defesa: SET NOCOUNT ON; evita mensagens de "linhas afetadas" que confundem o pytds
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
                
                # 2ª Defesa: Pular qualquer mensagem de status vazia até chegar na tabela real
                while cursor.description is None:
                    if not cursor.nextset():
                        break
                
                # Se achou uma tabela com colunas, extrai os dados
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
        
        # Tags exatas para Venda e Locação baseadas no fl_origem
        if 'fl_origem' in df.columns:
            df['origem_upper'] = df['fl_origem'].astype(str).str.strip().str.upper()
            tags_venda = ['VD', 'DV', 'IL', 'VENDA DE LOCAÇÃO']
            tags_locacao = ['FL', 'SL']
            df['is_venda'] = df['origem_upper'].isin(tags_venda)
            df['is_locacao'] = df['origem_upper'].isin(tags_locacao)
        else:
            df['is_venda'] = False
            df['is_locacao'] = False

        # Guardar os dados processados na memória antes de os devolver
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
    <!-- Atualização automática a cada 10 minutos (600s) -->
    <meta http-equiv="refresh" content="600">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mais Escoramentos | Faturamento</title>
    
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    
    <style>
        :root {
            --bg-color: #f8fafc; --card-bg: #ffffff; --text-main: #0f172a; --text-muted: #64748b;
            --primary: #0284c7; --success: #10b981; --warning: #f59e0b; --purple: #8b5cf6;
            --border-radius: 16px;
            --shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.05), 0 4px 6px -4px rgba(0, 0, 0, 0.05);
        }
        * { box-sizing: border-box; }
        body { font-family: 'Inter', sans-serif; margin: 0; padding: 20px 40px; background: var(--bg-color); color: var(--text-main); }
        
        #loader {
            display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(255,255,255,0.9); z-index: 9999; justify-content: center; align-items: center; flex-direction: column;
        }
        .spinner { border: 4px solid #e2e8f0; border-top: 4px solid var(--primary); border-radius: 50%; width: 50px; height: 50px; animation: spin 1s linear infinite; }
        @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }

        .header-container { display: flex; justify-content: space-between; align-items: center; margin-bottom: 35px; background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); padding: 25px 35px; border-radius: var(--border-radius); color: white; box-shadow: var(--shadow); }
        .header-title h1 { margin: 0; font-size: 26px; font-weight: 800; letter-spacing: -0.5px; }
        .header-title p { margin: 5px 0 0 0; font-size: 14px; color: #94a3b8; font-weight: 400; }

        .filtros-card { background: rgba(255, 255, 255, 0.1); padding: 12px 20px; border-radius: 12px; display: flex; align-items: center; gap: 15px; border: 1px solid rgba(255,255,255,0.2); backdrop-filter: blur(10px); }
        .filtros-card select { padding: 10px 18px; border-radius: 8px; border: none; background-color: white; font-family: 'Inter', sans-serif; font-weight: 600; color: var(--text-main); cursor: pointer; outline: none; }

        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 35px; }
        .kpi-card { background: var(--card-bg); padding: 25px; border-radius: var(--border-radius); box-shadow: var(--shadow); position: relative; overflow: hidden; display: flex; flex-direction: column; justify-content: center; transition: transform 0.3s ease; }
        .kpi-card:hover { transform: translateY(-5px); }
        .kpi-icon { position: absolute; right: -10px; bottom: -15px; font-size: 100px; opacity: 0.04; color: var(--text-main); transition: transform 0.3s ease; }
        .kpi-card:hover .kpi-icon { transform: scale(1.1) rotate(-5deg); }
        .kpi-icon-small { width: 40px; height: 40px; border-radius: 10px; display: flex; align-items: center; justify-content: center; margin-bottom: 15px; font-size: 18px; color: white; }
        .bg-primary { background: linear-gradient(135deg, #0ea5e9, #0284c7); } .bg-purple { background: linear-gradient(135deg, #a855f7, #7e22ce); } .bg-success { background: linear-gradient(135deg, #34d399, #059669); } .bg-warning { background: linear-gradient(135deg, #fbbf24, #d97706); }
        .kpi-title { font-size: 12px; font-weight: 700; text-transform: uppercase; color: var(--text-muted); margin-bottom: 5px; }
        .kpi-value { font-size: 28px; font-weight: 800; margin: 0; color: var(--text-main); letter-spacing: -1px; }
        .kpi-sub { font-size: 15px; font-weight: 700; margin: 0; color: var(--text-main); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }

        /* Novo layout de Gráficos (2/3 Mensal, 1/3 Anual) */
        .graficos-grid-layout { display: grid; grid-template-columns: 2fr 1fr; gap: 25px; margin-bottom: 25px; }
        .grafico-card { background: var(--card-bg); padding: 25px; border-radius: var(--border-radius); box-shadow: var(--shadow); height: 380px; display: flex; flex-direction: column; }
        .grafico-card-title { font-size: 16px; font-weight: 700; color: var(--text-main); margin-bottom: 15px; border-bottom: 1px solid #f1f5f9; padding-bottom: 10px; }
        
        .full-width-grid { display: grid; grid-template-columns: 1fr; margin-bottom: 25px; }
        .grafico-card-full { height: 450px; background: var(--card-bg); padding: 25px; border-radius: var(--border-radius); box-shadow: var(--shadow); }

        .status-footer { margin-top: 35px; text-align: center; font-size: 13px; color: var(--text-muted); display: flex; align-items: center; justify-content: center; gap: 10px; font-weight: 600; }
        .pulse { width: 10px; height: 10px; background-color: var(--success); border-radius: 50%; animation: pulse-animation 2s infinite; }
        @keyframes pulse-animation { 0% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.7); } 70% { box-shadow: 0 0 0 10px rgba(16, 185, 129, 0); } 100% { box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); } }

        @media (max-width: 1024px) { 
            .graficos-grid-layout { grid-template-columns: 1fr; } 
            .header-container { flex-direction: column; padding: 20px; gap: 20px; } 
            .filtros-card { width: 100%; justify-content: center; } 
        }
    </style>
</head>
<body>
    <div id="loader">
        <div class="spinner"></div>
        <h3 style="color: #0f172a; margin-top: 20px; font-family: 'Inter';">A atualizar os dados do BI...</h3>
    </div>

    <div class="header-container">
        <div class="header-title">
            <h1><i class="fa-solid fa-chart-line" style="margin-right: 10px; color: #38bdf8;"></i> Inteligência de Faturamento</h1>
            <p>Painel de Resultados • Mais Escoramentos</p>
        </div>
        
        <form method="GET" class="filtros-card" id="form-filtros">
            <span style="font-size: 13px; color: rgba(255,255,255,0.8); font-weight: 700; text-transform: uppercase;">Filtros:</span>
            <select name="ano" onchange="mostrarLoaderEEnviar()">
                {% for a in anos %}
                    <option value="{{ a }}" {% if a|string == ano_sel %}selected{% endif %}>Ano: {{ a }}</option>
                {% endfor %}
            </select>
            <select name="mes" onchange="mostrarLoaderEEnviar()">
                {% for v, n in meses %}
                    <option value="{{ v }}" {% if v|string == mes_sel %}selected{% endif %}>Mês: {{ n }}</option>
                {% endfor %}
            </select>
        </form>
    </div>

    <div class="kpi-grid">
        <div class="kpi-card">
            <i class="fa-solid fa-money-bill-trend-up kpi-icon"></i>
            <div class="kpi-icon-small bg-primary"><i class="fa-solid fa-dollar-sign"></i></div>
            <div class="kpi-title">Fat. Bruto (Total)</div>
            <div class="kpi-value">{{ total_faturamento }}</div>
        </div>
        <div class="kpi-card">
            <i class="fa-solid fa-tags kpi-icon"></i>
            <div class="kpi-icon-small bg-success"><i class="fa-solid fa-cart-shopping"></i></div>
            <div class="kpi-title">Vendas</div>
            <div class="kpi-value">{{ fat_vendas }}</div>
        </div>
        <div class="kpi-card">
            <i class="fa-solid fa-truck-fast kpi-icon"></i>
            <div class="kpi-icon-small" style="background: linear-gradient(135deg, #f43f5e, #e11d48);"><i class="fa-solid fa-dolly"></i></div>
            <div class="kpi-title">Locação</div>
            <div class="kpi-value">{{ fat_locacao }}</div>
        </div>
        <div class="kpi-card">
            <i class="fa-solid fa-chart-pie kpi-icon"></i>
            <div class="kpi-icon-small bg-warning"><i class="fa-solid fa-calculator"></i></div>
            <div class="kpi-title">Ticket Médio</div>
            <div class="kpi-value">{{ media }}</div>
        </div>
        <div class="kpi-card">
            <i class="fa-solid fa-building kpi-icon"></i>
            <div class="kpi-icon-small bg-purple"><i class="fa-solid fa-crown"></i></div>
            <div class="kpi-title">Top Cliente (R$)</div>
            <div class="kpi-sub" title="{{ top_c_nome }}">{{ top_c_nome }}</div>
        </div>
    </div>

    <!-- Seção: Locação -->
    <div class="graficos-grid-layout">
        <div class="grafico-card">
            <div class="grafico-card-title">Locação Mensal</div>
            <div id="g_loc_mes" style="width: 100%; flex-grow: 1;"></div>
        </div>
        <div class="grafico-card">
            <div class="grafico-card-title">Locação Anual</div>
            <div id="g_loc_ano" style="width: 100%; flex-grow: 1;"></div>
        </div>
    </div>

    <!-- Seção: Venda -->
    <div class="graficos-grid-layout">
        <div class="grafico-card">
            <div class="grafico-card-title">Venda Mensal</div>
            <div id="g_ven_mes" style="width: 100%; flex-grow: 1;"></div>
        </div>
        <div class="grafico-card">
            <div class="grafico-card-title">Venda Anual</div>
            <div id="g_ven_ano" style="width: 100%; flex-grow: 1;"></div>
        </div>
    </div>

    <!-- Seção: Top Clientes -->
    <div class="full-width-grid">
        <div class="grafico-card-full">
            <div class="grafico-card-title">Top 10 Clientes (Volume em R$)</div>
            <div id="g_top" style="width: 100%; height: 90%;"></div>
        </div>
    </div>

    <div class="status-footer">
        <div class="pulse"></div>
        <span>Nuvem 24h • Última extração: {{ data_extracao }} • Atualização em 10 min</span>
    </div>

    <script>
        function mostrarLoaderEEnviar() {
            document.getElementById('loader').style.display = 'flex';
            document.getElementById('form-filtros').submit();
        }

        var config = {responsive: true, displayModeBar: false};
        var g_loc_mes_data = {{ fig_loc_mes_json | safe }};
        var g_loc_ano_data = {{ fig_loc_ano_json | safe }};
        var g_ven_mes_data = {{ fig_ven_mes_json | safe }};
        var g_ven_ano_data = {{ fig_ven_ano_json | safe }};
        var g_top_data = {{ fig_top_json | safe }};
        
        Plotly.newPlot('g_loc_mes', g_loc_mes_data.data, g_loc_mes_data.layout, config);
        Plotly.newPlot('g_loc_ano', g_loc_ano_data.data, g_loc_ano_data.layout, config);
        Plotly.newPlot('g_ven_mes', g_ven_mes_data.data, g_ven_mes_data.layout, config);
        Plotly.newPlot('g_ven_ano', g_ven_ano_data.data, g_ven_ano_data.layout, config);
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
    # 1. CÁLCULOS DOS CARDS (Topos da tela) - Usa o filtro completo
    # -------------------------------------------------------------
    df_f = df.copy()
    if ano_sel != "all": df_f = df_f[df_f["ano"] == int(ano_sel)]
    if mes_sel != "all": df_f = df_f[df_f["mes_num"] == int(mes_sel)]

    total_faturamento = df_f["vl_faturamento_bruto"].sum()
    qtd_nfs = len(df_f)
    media = total_faturamento / qtd_nfs if qtd_nfs > 0 else 0
    
    faturamento_vendas = df_f[df_f['is_venda']]["vl_faturamento_bruto"].sum()
    faturamento_locacao = df_f[df_f['is_locacao']]["vl_faturamento_bruto"].sum()
        
    if not df_f.empty and total_faturamento > 0:
        top_c_nome = df_f.groupby("nm_cliente")["vl_faturamento_bruto"].sum().idxmax()
    else:
        top_c_nome = "Nenhum Cliente Registrado"

    # Layout unificado padrão do Plotly para evitar bugs de barras esmagadas
    layout_moderno = dict(
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Inter, sans-serif", color="#475569", size=12),
        margin=dict(l=10, r=10, t=10, b=10),
        hovermode="x unified"
    )

    meses_map = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}

    # -------------------------------------------------------------
    # 2. GRÁFICOS MENSAIS (Ignora filtro de mês, respeita o Ano)
    # -------------------------------------------------------------
    df_mensal_base = df.copy()
    if ano_sel != "all":
        df_mensal_base = df_mensal_base[df_mensal_base["ano"] == int(ano_sel)]

    # -> Locação Mensal
    df_loc_mes = df_mensal_base[df_mensal_base['is_locacao']].groupby('mes_num')["vl_faturamento_bruto"].sum().reset_index()
    df_loc_mes['mes_nome'] = df_loc_mes['mes_num'].map(meses_map)
    df_loc_mes = df_loc_mes.sort_values('mes_num')
    
    fig_loc_mes = px.bar(df_loc_mes, x="mes_nome", y="vl_faturamento_bruto")
    fig_loc_mes.update_traces(
        marker_color="#0ea5e9", # Azul
        text=df_loc_mes['vl_faturamento_bruto'].apply(formatar_k_m),
        textposition="auto" # "auto" previne que a barra desapareça
    )
    fig_loc_mes.update_layout(**layout_moderno)
    fig_loc_mes.update_yaxes(visible=False, showticklabels=False) # Esconde eixo Y para ficar limpo
    fig_loc_mes.update_xaxes(title="")

    # -> Venda Mensal
    df_ven_mes = df_mensal_base[df_mensal_base['is_venda']].groupby('mes_num')["vl_faturamento_bruto"].sum().reset_index()
    df_ven_mes['mes_nome'] = df_ven_mes['mes_num'].map(meses_map)
    df_ven_mes = df_ven_mes.sort_values('mes_num')
    
    fig_ven_mes = px.bar(df_ven_mes, x="mes_nome", y="vl_faturamento_bruto")
    fig_ven_mes.update_traces(
        marker_color="#10b981", # Verde
        text=df_ven_mes['vl_faturamento_bruto'].apply(formatar_k_m),
        textposition="auto"
    )
    fig_ven_mes.update_layout(**layout_moderno)
    fig_ven_mes.update_yaxes(visible=False, showticklabels=False)
    fig_ven_mes.update_xaxes(title="")

    # -------------------------------------------------------------
    # 3. GRÁFICOS ANUAIS (Ignora filtro de ano, respeita o Mês)
    # -------------------------------------------------------------
    df_anual_base = df.copy()
    if mes_sel != "all":
        df_anual_base = df_anual_base[df_anual_base["mes_num"] == int(mes_sel)]

    # -> Locação Anual
    df_loc_ano = df_anual_base[df_anual_base['is_locacao']].groupby('ano')["vl_faturamento_bruto"].sum().reset_index()
    fig_loc_ano = px.bar(df_loc_ano, x="ano", y="vl_faturamento_bruto")
    fig_loc_ano.update_traces(
        marker_color="#0ea5e9", # Azul
        text=df_loc_ano['vl_faturamento_bruto'].apply(formatar_k_m),
        textposition="auto"
    )
    fig_loc_ano.update_layout(**layout_moderno)
    fig_loc_ano.update_yaxes(visible=False, showticklabels=False)
    fig_loc_ano.update_xaxes(title="", type='category') # 'category' impede que o ano apareça quebrado ex: 2024.5

    # -> Venda Anual
    df_ven_ano = df_anual_base[df_anual_base['is_venda']].groupby('ano')["vl_faturamento_bruto"].sum().reset_index()
    fig_ven_ano = px.bar(df_ven_ano, x="ano", y="vl_faturamento_bruto")
    fig_ven_ano.update_traces(
        marker_color="#10b981", # Verde
        text=df_ven_ano['vl_faturamento_bruto'].apply(formatar_k_m),
        textposition="auto"
    )
    fig_ven_ano.update_layout(**layout_moderno)
    fig_ven_ano.update_yaxes(visible=False, showticklabels=False)
    fig_ven_ano.update_xaxes(title="", type='category')

    # -------------------------------------------------------------
    # 4. GRÁFICO TOP 10 CLIENTES (Usa o filtro completo)
    # -------------------------------------------------------------
    df_top = df_f.groupby("nm_cliente")["vl_faturamento_bruto"].sum().nlargest(10).reset_index()
    fig_top = px.bar(df_top, x="vl_faturamento_bruto", y="nm_cliente", orientation='h')
    fig_top.update_traces(
        marker_color="#8b5cf6", # Roxo
        text=df_top['vl_faturamento_bruto'].apply(formatar_k_m),
        textposition="auto"
    )
    fig_top.update_layout(yaxis={'categoryorder':'total ascending'}, **layout_moderno)
    fig_top.update_xaxes(visible=False, showticklabels=False)
    fig_top.update_yaxes(title="")

    # RECRIADO AQUI: Lista de meses para o Filtro HTML! 
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
        media=formatar_moeda(media),
        top_c_nome=str(top_c_nome)[:30], 
        fig_loc_mes_json=fig_loc_mes.to_json(),
        fig_loc_ano_json=fig_loc_ano.to_json(),
        fig_ven_mes_json=fig_ven_mes.to_json(),
        fig_ven_ano_json=fig_ven_ano.to_json(),
        fig_top_json=fig_top.to_json()
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
