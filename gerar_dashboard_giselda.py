import pandas as pd
import json
import shutil
import tempfile
import subprocess
import os
from pathlib import Path

def ler_dados_excel(caminho_arquivo):
    """Lê os dados do arquivo Excel do Hospital Giselda Trigueiro
    Se o arquivo estiver aberto, cria uma cópia temporária para leitura"""

    temp_file = Path(tempfile.gettempdir()) / "temp_custos_giselda.xlsx"

    # Tentar ler diretamente, se falhar criar cópia temporária via PowerShell
    try:
        df = pd.read_excel(caminho_arquivo, sheet_name=0, header=None)
    except PermissionError:
        print("   ⚠ Arquivo em uso pelo Excel, criando cópia temporária...")
        # Usar PowerShell para copiar (funciona mesmo com arquivo aberto)
        cmd = f'Copy-Item -Path "{caminho_arquivo}" -Destination "{temp_file}" -Force'
        subprocess.run(['powershell', '-Command', cmd], check=True, capture_output=True)
        df = pd.read_excel(temp_file, sheet_name=0, header=None)
        if temp_file.exists():
            temp_file.unlink()

    # Meses fixos (Janeiro a Dezembro de 2025)
    meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
    meses_completos = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                       'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

    # Função para extrair valores de uma linha
    def extrair_valores(linha_idx):
        valores = df.iloc[linha_idx, 1:13].tolist()
        resultado = []
        for v in valores:
            if pd.notna(v):
                try:
                    resultado.append(float(v))
                except (ValueError, TypeError):
                    resultado.append(0)
            else:
                resultado.append(0)
        return resultado

    # Dados principais (índices corretos)
    dados = {
        'meses': meses,
        'meses_completos': meses_completos,
        'pessoal': extrair_valores(7),        # Linha 7: Pessoal
        'material': extrair_valores(11),      # Linha 11: Material de Consumo
        'servicos': extrair_valores(33),      # Linha 33: Serviços de Terceiros
        'despesas': extrair_valores(55)       # Linha 55: Despesas Gerais
    }

    # Subcategorias de Material de Consumo
    dados['materiais_subcategorias'] = {
        'Medicamentos': extrair_valores(27),
        'Gêneros de Alimentação': extrair_valores(16),
        'Vacinas': extrair_valores(31),
        'Nutrição Enteral': extrair_valores(28),
        'Gases Medicinais': extrair_valores(14),
        'Material de Limpeza': extrair_valores(21),
        'Fórmulas Nutricionais': extrair_valores(13),
        'GLP': extrair_valores(15),
        'Embalagem': extrair_valores(17),
        'Combustíveis': extrair_valores(12),
        'Material de Expediente': extrair_valores(20),
        'Material Laboratorial': extrair_valores(23)
    }

    # Subcategorias de Serviços de Terceiros
    dados['servicos_subcategorias'] = {
        'Limpeza e Conservação': extrair_valores(50),
        'Vigilância/Segurança': extrair_valores(52),
        'Apoio Administrativo': extrair_valores(47),
        'Médico-Hospitalares': extrair_valores(53),
        'Manutenção Bens Imóveis': extrair_valores(42),
        'Lavanderia': extrair_valores(40),
        'Manutenção Máquinas': extrair_valores(43),
        'Tratamento de Resíduos': extrair_valores(46),
        'Laboratório': extrair_valores(36),
        'Transporte': extrair_valores(45),
        'Aluguel Máquinas': extrair_valores(34),
        'Aluguel Veículos': extrair_valores(35)
    }

    return dados

def gerar_html(dados):
    """Gera o HTML do dashboard com os dados"""

    # Calcular totais mensais
    totais_mensais = [
        dados['pessoal'][i] + dados['material'][i] + dados['servicos'][i] + dados['despesas'][i]
        for i in range(len(dados['meses']))
    ]

    html = f'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard - Hospital Giselda Trigueiro 2025</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}

        body {{
            background: linear-gradient(135deg, #f5f7fa 0%, #e4e8f0 100%);
            min-height: 100vh;
            padding: 0;
            display: flex;
        }}

        .sidebar {{
            width: 280px;
            background: #15325b;
            color: white;
            padding: 25px 20px;
            height: 100vh;
            position: sticky;
            top: 0;
            overflow-y: auto;
            box-shadow: 2px 0 10px rgba(0,0,0,0.1);
        }}

        .sidebar h2 {{
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 25px;
            color: #fff;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}

        .filter-group {{
            margin-bottom: 20px;
        }}

        .filter-group label {{
            display: block;
            color: #94a3b8;
            font-size: 12px;
            font-weight: 600;
            text-transform: uppercase;
            margin-bottom: 8px;
        }}

        .filter-group select {{
            width: 100%;
            padding: 10px;
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: 6px;
            background: rgba(255, 255, 255, 0.1);
            color: white;
            font-size: 14px;
            outline: none;
            cursor: pointer;
        }}

        .filter-group select option {{
            background: #1e3a5f;
            color: white;
        }}

        .btn-clear {{
            width: 100%;
            padding: 12px;
            background: rgba(255, 255, 255, 0.1);
            border: 1px solid rgba(255, 255, 255, 0.2);
            color: white;
            border-radius: 6px;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.3s;
            margin-top: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
        }}

        .btn-clear:hover {{
            background: rgba(255, 255, 255, 0.2);
            border-color: rgba(255, 255, 255, 0.4);
        }}

        .main-content {{
            flex: 1;
            padding: 20px;
            max-width: calc(100% - 280px);
            overflow-y: auto;
        }}

        .header {{
            background: linear-gradient(135deg, #1e3a5f 0%, #2563eb 100%);
            color: white;
            padding: 25px 30px;
            border-radius: 15px;
            margin-bottom: 25px;
            box-shadow: 0 4px 15px rgba(37, 99, 235, 0.3);
        }}

        .header h1 {{
            font-size: 28px;
            font-weight: 600;
            margin-bottom: 5px;
        }}

        .header p {{
            opacity: 0.9;
            font-size: 14px;
        }}

        .kpi-container {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 20px;
            margin-bottom: 25px;
        }}

        .kpi-card {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
            border-left: 4px solid #2563eb;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}

        .kpi-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 8px 25px rgba(37, 99, 235, 0.2);
        }}

        .kpi-card h3 {{
            color: #64748b;
            font-size: 13px;
            font-weight: 500;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            margin-bottom: 10px;
        }}

        .kpi-card .value {{
            color: #1e3a5f;
            font-size: 28px;
            font-weight: 700;
        }}

        .kpi-card .subtext {{
            color: #94a3b8;
            font-size: 12px;
            margin-top: 8px;
        }}

        .kpi-card.highlight {{
            border-left-color: #059669;
        }}

        .kpi-card.warning {{
            border-left-color: #f59e0b;
        }}

        .kpi-card.danger {{
            border-left-color: #ef4444;
        }}

        .charts-container {{
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 20px;
            margin-bottom: 25px;
        }}

        .chart-card {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
        }}

        .chart-card.full-width {{
            grid-column: 1 / -1;
        }}

        .chart-card h3 {{
            color: #1e3a5f;
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #e2e8f0;
        }}

        .chart-wrapper {{
            position: relative;
            height: 300px;
        }}

        canvas {{
            cursor: pointer;
        }}

        .chart-card {{
            transition: box-shadow 0.2s ease;
        }}

        .chart-wrapper.large {{
            height: 400px;
        }}

        .table-container {{
            background: white;
            border-radius: 15px;
            padding: 25px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.08);
            overflow-x: auto;
        }}

        .table-container h3 {{
            color: #1e3a5f;
            font-size: 16px;
            font-weight: 600;
            margin-bottom: 20px;
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
        }}

        th, td {{
            padding: 12px 15px;
            text-align: left;
            border-bottom: 1px solid #e2e8f0;
        }}

        th {{
            background: #f8fafc;
            color: #1e3a5f;
            font-weight: 600;
            font-size: 13px;
        }}

        td {{
            color: #475569;
            font-size: 14px;
        }}

        tr:hover {{
            background: #f8fafc;
        }}

        .positive {{
            color: #059669;
        }}

        .negative {{
            color: #ef4444;
        }}

        @media (max-width: 768px) {{
            body {{
                flex-direction: column;
            }}
            .sidebar {{
                width: 100%;
                height: auto;
                position: relative;
            }}
            .main-content {{
                max-width: 100%;
            }}
            .charts-container {{
                grid-template-columns: 1fr;
            }}

            .header h1 {{
                font-size: 22px;
            }}

            .kpi-card .value {{
                font-size: 22px;
            }}
        }}
    </style>
</head>
<body>
    <div class="sidebar">
        <h2>Filtros do Dashboard</h2>
        
        <div class="filter-group">
            <label>Filtrar por Categoria:</label>
            <select id="categoryFilter" onchange="filterByCategory()">
                <option value="all">Todas as Categorias</option>
                <option value="pessoal">Pessoal</option>
                <option value="material">Material de Consumo</option>
                <option value="servicos">Serviços de Terceiros</option>
                <option value="despesas">Despesas Gerais</option>
            </select>
        </div>

        <div class="filter-group">
            <label>Período:</label>
            <select id="periodFilter" onchange="filterByPeriod()">
                <option value="all">Ano Completo</option>
                <option value="q1">1º Trimestre</option>
                <option value="q2">2º Trimestre</option>
                <option value="q3">3º Trimestre</option>
                <option value="q4">4º Trimestre</option>
            </select>
        </div>

        <button class="btn-clear" onclick="limparFiltros()">
            ✕ Limpar Filtros
        </button>

        <div id="activeFiltersBadge" style="display:none; margin-top:14px; background:rgba(37,99,235,0.18); border:1px solid rgba(96,165,250,0.45); border-radius:8px; padding:10px 12px; font-size:12px; color:#93c5fd; line-height:1.7;"></div>

        <p style="color:#475569; font-size:11px; margin-top:18px; text-align:center; line-height:1.5;">💡 Clique nos gráficos para filtrar.<br>Clique novamente para desfazer.</p>
    </div>

    <div class="main-content">
        <div class="header">
            <h1>Hospital Giselda Trigueiro</h1>
            <p>Dashboard de Custos - Ano 2025 | Gerado automaticamente via Python</p>
        </div>

        <div class="kpi-container">
        <div class="kpi-card">
            <h3>Custo Total Anual</h3>
            <div class="value" id="kpiTotal">R$ 0</div>
            <div class="subtext">Acumulado 2025</div>
        </div>
        <div class="kpi-card highlight">
            <h3>Média Mensal</h3>
            <div class="value" id="kpiMedia">R$ 0</div>
            <div class="subtext">Média dos 12 meses</div>
        </div>
        <div class="kpi-card warning">
            <h3>Maior Custo</h3>
            <div class="value" id="kpiMaior">R$ 0</div>
            <div class="subtext" id="kpiMaiorMes">-</div>
        </div>
        <div class="kpi-card danger">
            <h3>Menor Custo</h3>
            <div class="value" id="kpiMenor">R$ 0</div>
            <div class="subtext" id="kpiMenorMes">-</div>
        </div>
    </div>

    <div class="charts-container">
        <div class="chart-card full-width">
            <h3>Evolução Mensal dos Custos por Categoria</h3>
            <div class="chart-wrapper large">
                <canvas id="lineChart"></canvas>
            </div>
        </div>

        <div class="chart-card">
            <h3>Distribuição por Categoria</h3>
            <div class="chart-wrapper">
                <canvas id="pieChart"></canvas>
            </div>
        </div>

        <div class="chart-card">
            <h3>Comparativo Mensal - Custo Total</h3>
            <div class="chart-wrapper">
                <canvas id="barChart"></canvas>
            </div>
        </div>

        <div class="chart-card">
            <h3>Top Materiais de Consumo</h3>
            <div class="chart-wrapper">
                <canvas id="horizontalBarChart"></canvas>
            </div>
        </div>

        <div class="chart-card">
            <h3>Serviços de Terceiros - Detalhamento</h3>
            <div class="chart-wrapper">
                <canvas id="doughnutChart"></canvas>
            </div>
        </div>
    </div>

    <div class="table-container">
        <h3>Resumo Mensal por Categoria</h3>
        <table id="summaryTable">
            <thead>
                <tr>
                    <th>Mês</th>
                    <th>Pessoal</th>
                    <th>Material</th>
                    <th>Serviços</th>
                    <th>Despesas</th>
                    <th>Total</th>
                    <th>Variação</th>
                </tr>
            </thead>
            <tbody id="tableBody">
            </tbody>
        </table>
    </div>

    <script>
        // Dados carregados do Excel via Python
        const meses = {json.dumps(dados['meses'])};
        const mesesCompletos = {json.dumps(dados['meses_completos'])};
        const dadosPessoal = {json.dumps(dados['pessoal'])};
        const dadosMaterial = {json.dumps(dados['material'])};
        const dadosServicos = {json.dumps(dados['servicos'])};
        const dadosDespesas = {json.dumps(dados['despesas'])};
        const materiaisSubcategorias = {json.dumps(dados['materiais_subcategorias'])};
        const servicosSubcategorias = {json.dumps(dados['servicos_subcategorias'])};

        // Calcular totais mensais
        const totaisMensais = dadosPessoal.map((val, idx) =>
            val + dadosMaterial[idx] + dadosServicos[idx] + dadosDespesas[idx]
        );

        // Variáveis para os gráficos
        let lineChart, pieChart, barChart, horizontalBarChart, doughnutChart;
        let activeCategories = ['pessoal', 'material', 'servicos', 'despesas'];
        let currentPeriod = 'all';
        let selectedMonth = null;
        let selectedCategory = null;
        let selectedMaterial = null;

        // Cores do tema
        const cores = {{
            pessoal: {{ bg: 'rgba(37, 99, 235, 0.8)', border: 'rgb(37, 99, 235)' }},
            material: {{ bg: 'rgba(5, 150, 105, 0.8)', border: 'rgb(5, 150, 105)' }},
            servicos: {{ bg: 'rgba(245, 158, 11, 0.8)', border: 'rgb(245, 158, 11)' }},
            despesas: {{ bg: 'rgba(239, 68, 68, 0.8)', border: 'rgb(239, 68, 68)' }}
        }};

        const coresArray = [
            'rgba(37, 99, 235, 0.85)',   // Azul
            'rgba(16, 185, 129, 0.85)',  // Esmeralda
            'rgba(245, 158, 11, 0.85)',  // Âmbar
            'rgba(239, 68, 68, 0.85)',   // Vermelho
            'rgba(139, 92, 246, 0.85)',  // Violeta
            'rgba(20, 184, 166, 0.85)',  // Teal
            'rgba(236, 72, 153, 0.85)',  // Rosa
            'rgba(234, 179, 8, 0.85)',   // Amarelo
            'rgba(99, 102, 241, 0.85)',  // Índigo
            'rgba(6, 182, 212, 0.85)',   // Ciano
            'rgba(132, 204, 22, 0.85)',  // Lima
            'rgba(244, 63, 94, 0.85)'    // Rose
        ];

        const coresBordaArray = [
            'rgb(37, 99, 235)',
            'rgb(16, 185, 129)',
            'rgb(245, 158, 11)',
            'rgb(239, 68, 68)',
            'rgb(139, 92, 246)',
            'rgb(20, 184, 166)',
            'rgb(236, 72, 153)',
            'rgb(234, 179, 8)',
            'rgb(99, 102, 241)',
            'rgb(6, 182, 212)',
            'rgb(132, 204, 22)',
            'rgb(244, 63, 94)'
        ];

        // Funções utilitárias
        function formatCurrency(value) {{
            return new Intl.NumberFormat('pt-BR', {{
                style: 'currency',
                currency: 'BRL',
                minimumFractionDigits: 0,
                maximumFractionDigits: 0
            }}).format(value);
        }}

        function formatCurrencyShort(value) {{
            if (value >= 1000000) {{
                return 'R$ ' + (value / 1000000).toFixed(1) + 'M';
            }} else if (value >= 1000) {{
                return 'R$ ' + (value / 1000).toFixed(0) + 'K';
            }}
            return formatCurrency(value);
        }}

        function getDataByPeriod(data, period) {{
            switch(period) {{
                case 'q1': return data.slice(0, 3);
                case 'q2': return data.slice(3, 6);
                case 'q3': return data.slice(6, 9);
                case 'q4': return data.slice(9, 12);
                default: return data;
            }}
        }}

        function getMesesByPeriod(period) {{
            switch(period) {{
                case 'q1': return meses.slice(0, 3);
                case 'q2': return meses.slice(3, 6);
                case 'q3': return meses.slice(6, 9);
                case 'q4': return meses.slice(9, 12);
                default: return meses;
            }}
        }}

        // Atualizar KPIs
        function updateKPIs() {{
            const period = document.getElementById('periodFilter').value;
            const totais = getDataByPeriod(totaisMensais, period);
            const mesesPeriodo = period === 'all' ? mesesCompletos : getMesesByPeriod(period).map(m => {{
                const idx = meses.indexOf(m);
                return mesesCompletos[idx];
            }});

            const total = totais.reduce((a, b) => a + b, 0);
            const media = total / totais.length;
            const maior = Math.max(...totais);
            const menor = Math.min(...totais);
            const idxMaior = totais.indexOf(maior);
            const idxMenor = totais.indexOf(menor);

            document.getElementById('kpiTotal').textContent = formatCurrency(total);
            document.getElementById('kpiMedia').textContent = formatCurrency(media);
            document.getElementById('kpiMaior').textContent = formatCurrency(maior);
            document.getElementById('kpiMaiorMes').textContent = mesesPeriodo[idxMaior];
            document.getElementById('kpiMenor').textContent = formatCurrency(menor);
            document.getElementById('kpiMenorMes').textContent = mesesPeriodo[idxMenor];
        }}

        // Criar gráfico de linha
        function createLineChart() {{
            const ctx = document.getElementById('lineChart').getContext('2d');
            const period = document.getElementById('periodFilter').value;

            const datasets = [];

            if (activeCategories.includes('pessoal')) {{
                datasets.push({{
                    label: 'Pessoal',
                    data: getDataByPeriod(dadosPessoal, period),
                    borderColor: cores.pessoal.border,
                    backgroundColor: cores.pessoal.bg,
                    tension: 0.4,
                    fill: false,
                    pointRadius: 5,
                    pointHoverRadius: 8
                }});
            }}

            if (activeCategories.includes('material')) {{
                datasets.push({{
                    label: 'Material de Consumo',
                    data: getDataByPeriod(dadosMaterial, period),
                    borderColor: cores.material.border,
                    backgroundColor: cores.material.bg,
                    tension: 0.4,
                    fill: false,
                    pointRadius: 5,
                    pointHoverRadius: 8
                }});
            }}

            if (activeCategories.includes('servicos')) {{
                datasets.push({{
                    label: 'Serviços de Terceiros',
                    data: getDataByPeriod(dadosServicos, period),
                    borderColor: cores.servicos.border,
                    backgroundColor: cores.servicos.bg,
                    tension: 0.4,
                    fill: false,
                    pointRadius: 5,
                    pointHoverRadius: 8
                }});
            }}

            if (activeCategories.includes('despesas')) {{
                datasets.push({{
                    label: 'Despesas Gerais',
                    data: getDataByPeriod(dadosDespesas, period),
                    borderColor: cores.despesas.border,
                    backgroundColor: cores.despesas.bg,
                    tension: 0.4,
                    fill: false,
                    pointRadius: 5,
                    pointHoverRadius: 8
                }});
            }}

            if (lineChart) lineChart.destroy();

            lineChart = new Chart(ctx, {{
                type: 'line',
                data: {{
                    labels: getMesesByPeriod(period),
                    datasets: datasets
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    interaction: {{
                        mode: 'index',
                        intersect: false
                    }},
                    plugins: {{
                        legend: {{
                            position: 'top'
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return context.dataset.label + ': ' + formatCurrency(context.raw);
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            ticks: {{
                                callback: function(value) {{
                                    return formatCurrencyShort(value);
                                }}
                            }}
                        }}
                    }},
                    onClick: function(evt, elements) {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            highlightMonth(idx);
                        }}
                    }}
                }}
            }});
        }}

        // Criar gráfico de pizza
        function createPieChart() {{
            const ctx = document.getElementById('pieChart').getContext('2d');
            const period = document.getElementById('periodFilter').value;

            const totals = [
                getDataByPeriod(dadosPessoal, period).reduce((a, b) => a + b, 0),
                getDataByPeriod(dadosMaterial, period).reduce((a, b) => a + b, 0),
                getDataByPeriod(dadosServicos, period).reduce((a, b) => a + b, 0),
                getDataByPeriod(dadosDespesas, period).reduce((a, b) => a + b, 0)
            ];

            if (pieChart) pieChart.destroy();

            pieChart = new Chart(ctx, {{
                type: 'pie',
                data: {{
                    labels: ['Pessoal', 'Material de Consumo', 'Serviços de Terceiros', 'Despesas Gerais'],
                    datasets: [{{
                        data: totals,
                        backgroundColor: [
                            cores.pessoal.bg,
                            cores.material.bg,
                            cores.servicos.bg,
                            cores.despesas.bg
                        ],
                        borderColor: [
                            cores.pessoal.border,
                            cores.material.border,
                            cores.servicos.border,
                            cores.despesas.border
                        ],
                        borderWidth: 2
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'bottom'
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = ((context.raw / total) * 100).toFixed(1);
                                    return context.label + ': ' + formatCurrency(context.raw) + ' (' + percentage + '%)';
                                }}
                            }}
                        }}
                    }},
                    onClick: function(evt, elements) {{
                        if (elements.length > 0) {{
                            const categories = ['pessoal', 'material', 'servicos', 'despesas'];
                            const category = categories[elements[0].index];
                            if (selectedCategory === category) {{
                                selectedCategory = null;
                                activeCategories = ['pessoal', 'material', 'servicos', 'despesas'];
                                document.getElementById('categoryFilter').value = 'all';
                            }} else {{
                                selectedCategory = category;
                                activeCategories = [category];
                                document.getElementById('categoryFilter').value = category;
                            }}
                            updateActiveFiltersDisplay();
                            createLineChart();
                            updateKPIs();
                        }}
                    }}
                }}
            }});
        }}

        // Criar gráfico de barras
        function createBarChart() {{
            const ctx = document.getElementById('barChart').getContext('2d');
            const period = document.getElementById('periodFilter').value;

            if (barChart) barChart.destroy();

            barChart = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: getMesesByPeriod(period),
                    datasets: [{{
                        label: 'Custo Total',
                        data: getDataByPeriod(totaisMensais, period),
                        backgroundColor: getDataByPeriod(totaisMensais, period).map((val, idx) =>
                            selectedMonth !== null && idx !== selectedMonth
                                ? 'rgba(37, 99, 235, 0.18)'
                                : 'rgba(37, 99, 235, 0.85)'
                        ),
                        borderColor: getDataByPeriod(totaisMensais, period).map((val, idx) =>
                            selectedMonth !== null && idx !== selectedMonth
                                ? 'rgba(37, 99, 235, 0.30)'
                                : 'rgb(37, 99, 235)'
                        ),
                        borderWidth: 2
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            display: false
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return 'Total: ' + formatCurrency(context.raw);
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        y: {{
                            beginAtZero: true,
                            ticks: {{
                                callback: function(value) {{
                                    return formatCurrencyShort(value);
                                }}
                            }}
                        }}
                    }},
                    onClick: function(evt, elements) {{
                        if (elements.length > 0) {{
                            const idx = elements[0].index;
                            if (selectedMonth === idx) {{
                                selectedMonth = null;
                                resetHighlights();
                            }} else {{
                                selectedMonth = idx;
                                highlightMonth(idx);
                            }}
                            updateActiveFiltersDisplay();
                        }} else {{
                            selectedMonth = null;
                            resetHighlights();
                            updateActiveFiltersDisplay();
                        }}
                    }}
                }}
            }});
        }}

        // Criar gráfico de barras horizontais (Top Materiais)
        function createHorizontalBarChart() {{
            const ctx = document.getElementById('horizontalBarChart').getContext('2d');
            const period = document.getElementById('periodFilter').value;

            const totals = {{}};
            for (const [key, values] of Object.entries(materiaisSubcategorias)) {{
                totals[key] = getDataByPeriod(values, period).reduce((a, b) => a + b, 0);
            }}

            const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]);

            if (horizontalBarChart) horizontalBarChart.destroy();

            horizontalBarChart = new Chart(ctx, {{
                type: 'bar',
                data: {{
                    labels: sorted.map(item => item[0]),
                    datasets: [{{
                        label: 'Total',
                        data: sorted.map(item => item[1]),
                        backgroundColor: sorted.map(item =>
                            selectedMaterial !== null && item[0] !== selectedMaterial
                                ? 'rgba(203, 213, 225, 0.4)'
                                : coresArray[sorted.indexOf(item) % coresArray.length]
                        ),
                        borderColor: sorted.map(item =>
                            selectedMaterial !== null && item[0] !== selectedMaterial
                                ? 'rgba(203, 213, 225, 0.6)'
                                : coresBordaArray[sorted.indexOf(item) % coresBordaArray.length]
                        ),
                        borderWidth: 1,
                        borderRadius: 4
                    }}]
                }},
                options: {{
                    indexAxis: 'y',
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            display: false
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    return formatCurrency(context.raw);
                                }}
                            }}
                        }}
                    }},
                    scales: {{
                        x: {{
                            beginAtZero: true,
                            ticks: {{
                                callback: function(value) {{
                                    return formatCurrencyShort(value);
                                }}
                            }}
                        }}
                    }},
                    onClick: function(evt, elements) {{
                        if (elements.length > 0) {{
                            const label = sorted[elements[0].index][0];
                            if (selectedMaterial === label) {{
                                selectedMaterial = null;
                            }} else {{
                                selectedMaterial = label;
                            }}
                            updateActiveFiltersDisplay();
                            createHorizontalBarChart();
                        }} else {{
                            selectedMaterial = null;
                            updateActiveFiltersDisplay();
                            createHorizontalBarChart();
                        }}
                    }}
                }}
            }});
        }}

        // Criar gráfico de rosca (Serviços)
        function createDoughnutChart() {{
            const ctx = document.getElementById('doughnutChart').getContext('2d');
            const period = document.getElementById('periodFilter').value;

            const totals = {{}};
            for (const [key, values] of Object.entries(servicosSubcategorias)) {{
                totals[key] = getDataByPeriod(values, period).reduce((a, b) => a + b, 0);
            }}

            const sorted = Object.entries(totals).sort((a, b) => b[1] - a[1]);

            if (doughnutChart) doughnutChart.destroy();

            doughnutChart = new Chart(ctx, {{
                type: 'doughnut',
                data: {{
                    labels: sorted.map(item => item[0]),
                    datasets: [{{
                        data: sorted.map(item => item[1]),
                        backgroundColor: coresArray,
                        borderColor: '#fff',
                        borderWidth: 2
                    }}]
                }},
                options: {{
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {{
                        legend: {{
                            position: 'right',
                            labels: {{
                                boxWidth: 12,
                                font: {{
                                    size: 11
                                }}
                            }}
                        }},
                        tooltip: {{
                            callbacks: {{
                                label: function(context) {{
                                    const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                    const percentage = ((context.raw / total) * 100).toFixed(1);
                                    return context.label + ': ' + formatCurrency(context.raw) + ' (' + percentage + '%)';
                                }}
                            }}
                        }}
                    }}
                }}
            }});
        }}

        // Criar tabela resumo
        function createTable() {{
            const tbody = document.getElementById('tableBody');
            tbody.innerHTML = '';

            const period = document.getElementById('periodFilter').value;
            const mesesFiltrados = getMesesByPeriod(period);
            const startIdx = period === 'all' ? 0 :
                             period === 'q1' ? 0 :
                             period === 'q2' ? 3 :
                             period === 'q3' ? 6 : 9;

            mesesFiltrados.forEach((mes, idx) => {{
                const realIdx = startIdx + idx;
                const pessoal = dadosPessoal[realIdx];
                const material = dadosMaterial[realIdx];
                const servicos = dadosServicos[realIdx];
                const despesas = dadosDespesas[realIdx];
                const total = totaisMensais[realIdx];

                let variacao = 0;
                let variacaoClass = '';
                if (realIdx > 0) {{
                    variacao = ((total - totaisMensais[realIdx - 1]) / totaisMensais[realIdx - 1]) * 100;
                    variacaoClass = variacao >= 0 ? 'negative' : 'positive';
                }}

                const row = document.createElement('tr');
                row.innerHTML = `
                    <td><strong>${{mesesCompletos[realIdx]}}</strong></td>
                    <td>${{formatCurrency(pessoal)}}</td>
                    <td>${{formatCurrency(material)}}</td>
                    <td>${{formatCurrency(servicos)}}</td>
                    <td>${{formatCurrency(despesas)}}</td>
                    <td><strong>${{formatCurrency(total)}}</strong></td>
                    <td class="${{variacaoClass}}">${{realIdx > 0 ? (variacao >= 0 ? '+' : '') + variacao.toFixed(1) + '%' : '-'}}</td>
                `;
                tbody.appendChild(row);
            }});
        }}

        // Filtrar por categoria
        function filterByCategory() {{
            const category = document.getElementById('categoryFilter').value;

            if (category === 'all') {{
                activeCategories = ['pessoal', 'material', 'servicos', 'despesas'];
            }} else {{
                activeCategories = [category];
            }}

            updateAllCharts();
        }}

        // Filtrar por período
        function filterByPeriod() {{
            currentPeriod = document.getElementById('periodFilter').value;
            updateAllCharts();
        }}

        // Limpar todos os filtros
        function limparFiltros() {{
            document.getElementById('categoryFilter').value = 'all';
            document.getElementById('periodFilter').value = 'all';
            activeCategories = ['pessoal', 'material', 'servicos', 'despesas'];
            currentPeriod = 'all';
            selectedMonth = null;
            selectedCategory = null;
            selectedMaterial = null;
            updateActiveFiltersDisplay();
            updateAllCharts();
        }}

        // Destacar mês em todos os gráficos (cross-filtering estilo Power BI)
        function highlightMonth(monthIdx) {{
            const period = document.getElementById('periodFilter').value;
            const startIdx = period === 'q2' ? 3 : period === 'q3' ? 6 : period === 'q4' ? 9 : 0;
            const realIdx = startIdx + monthIdx;
            const n = barChart ? barChart.data.datasets[0].data.length : 12;

            // Destacar barra selecionada, dimming nas outras
            if (barChart) {{
                barChart.data.datasets[0].backgroundColor = Array.from({{length: n}}, (_, i) =>
                    i === monthIdx ? 'rgba(37, 99, 235, 1)' : 'rgba(37, 99, 235, 0.18)'
                );
                barChart.data.datasets[0].borderColor = Array.from({{length: n}}, (_, i) =>
                    i === monthIdx ? 'rgb(14, 60, 180)' : 'rgba(37, 99, 235, 0.30)'
                );
                barChart.update('none');
            }}

            // Destacar ponto no gráfico de linha
            if (lineChart) {{
                lineChart.data.datasets.forEach(ds => {{
                    ds.pointRadius = ds.data.map((_, i) => i === monthIdx ? 10 : 3);
                    ds.pointHoverRadius = ds.data.map((_, i) => i === monthIdx ? 13 : 6);
                    ds.pointBorderWidth = ds.data.map((_, i) => i === monthIdx ? 3 : 1);
                }});
                lineChart.update('none');
            }}
        }}

        // Resetar todos os destaques visuais
        function resetHighlights() {{
            if (barChart) {{
                const n = barChart.data.datasets[0].data.length;
                barChart.data.datasets[0].backgroundColor = Array.from({{length: n}}, () => 'rgba(37, 99, 235, 0.85)');
                barChart.data.datasets[0].borderColor = Array.from({{length: n}}, () => 'rgb(37, 99, 235)');
                barChart.update('none');
            }}
            if (lineChart) {{
                lineChart.data.datasets.forEach(ds => {{
                    ds.pointRadius = 5;
                    ds.pointHoverRadius = 8;
                    ds.pointBorderWidth = 1;
                }});
                lineChart.update('none');
            }}
        }}

        // Exibir filtros ativos no sidebar
        function updateActiveFiltersDisplay() {{
            const badge = document.getElementById('activeFiltersBadge');
            if (!badge) return;
            const parts = [];
            if (selectedMonth !== null) {{
                const period = document.getElementById('periodFilter').value;
                const ml = getMesesByPeriod(period);
                parts.push('📅 ' + (ml[selectedMonth] || ''));
            }}
            if (selectedCategory !== null) {{
                const names = {{ pessoal: 'Pessoal', material: 'Material', servicos: 'Serviços', despesas: 'Despesas' }};
                parts.push('📂 ' + names[selectedCategory]);
            }}
            if (selectedMaterial !== null) {{
                parts.push('📦 ' + selectedMaterial);
            }}
            const periodVal = document.getElementById('periodFilter').value;
            if (periodVal !== 'all') {{
                const pn = {{ q1: '1º Trim', q2: '2º Trim', q3: '3º Trim', q4: '4º Trim' }};
                parts.push('🗓 ' + pn[periodVal]);
            }}
            if (parts.length > 0) {{
                badge.innerHTML = '<strong>Ativos:</strong><br>' + parts.join('<br>');
                badge.style.display = 'block';
            }} else {{
                badge.style.display = 'none';
            }}
        }}

        // Atualizar todos os gráficos
        function updateAllCharts() {{
            updateKPIs();
            createLineChart();
            createPieChart();
            createBarChart();
            createHorizontalBarChart();
            createDoughnutChart();
            createTable();
        }}

        // Inicializar
        document.addEventListener('DOMContentLoaded', function() {{
            updateAllCharts();
        }});
    </script>
    </div>
</body>
</html>'''

    return html

def main():
    # Caminho do arquivo Excel
    pasta = Path(__file__).parent
    arquivo_excel = pasta / "Custo_Total_da_Unidade Giselda Trigueiro 2025.xlsx"
    arquivo_html = pasta / "dashboard_hospital_giselda.html"

    print("=" * 60)
    print("GERADOR DE DASHBOARD - HOSPITAL GISELDA TRIGUEIRO")
    print("=" * 60)

    # Verificar se o arquivo existe
    if not arquivo_excel.exists():
        print(f"\n❌ Erro: Arquivo não encontrado!")
        print(f"   Procurando em: {arquivo_excel}")
        return

    print(f"\n📊 Lendo dados do Excel...")
    print(f"   Arquivo: {arquivo_excel.name}")

    try:
        # Ler dados
        dados = ler_dados_excel(arquivo_excel)
        print(f"   ✓ {len(dados['meses'])} meses carregados")
        print(f"   ✓ {len(dados['materiais_subcategorias'])} subcategorias de materiais")
        print(f"   ✓ {len(dados['servicos_subcategorias'])} subcategorias de serviços")

        # Gerar HTML
        print(f"\n🔧 Gerando dashboard HTML...")
        html = gerar_html(dados)

        # Salvar arquivo
        with open(arquivo_html, 'w', encoding='utf-8') as f:
            f.write(html)

        print(f"   ✓ Dashboard salvo em: {arquivo_html.name}")

        # Calcular totais para exibir
        total_anual = sum(
            dados['pessoal'][i] + dados['material'][i] + dados['servicos'][i] + dados['despesas'][i]
            for i in range(len(dados['meses']))
        )

        print(f"\n📈 Resumo dos dados:")
        print(f"   Custo Total Anual: R$ {total_anual:,.2f}")
        print(f"   Média Mensal: R$ {total_anual/12:,.2f}")

        print(f"\n✅ Dashboard gerado com sucesso!")
        print(f"\n💡 Para visualizar, abra o arquivo no navegador:")
        print(f"   {arquivo_html}")

    except Exception as e:
        print(f"\n❌ Erro ao processar: {e}")
        raise

if __name__ == "__main__":
    main()
