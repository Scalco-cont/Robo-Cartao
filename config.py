"""
Arquivo de configuração do Processador de Cartões
Centralize todas as configurações aqui para facilitar manutenção
"""

import os

# =========================
# CONFIGURAÇÕES GERAIS
# =========================
class Config:
    # Flask
    SECRET_KEY = os.environ.get('SECRET_KEY') or os.urandom(24)
    DEBUG = os.environ.get('FLASK_DEBUG', 'True') == 'True'
    HOST = os.environ.get('FLASK_HOST', '0.0.0.0')
    PORT = int(os.environ.get('FLASK_PORT', 2929))
    
    # Pastas
    UPLOAD_FOLDER = os.environ.get('UPLOAD_FOLDER', 'uploads')
    OUTPUT_FOLDER = os.environ.get('OUTPUT_FOLDER', 'outputs')
    
    # Extensões permitidas
    ALLOWED_EXTENSIONS = {'xlsx'}
    
    # Tamanho máximo de arquivo (em bytes) - 16MB padrão
    MAX_FILE_SIZE = int(os.environ.get('MAX_FILE_SIZE', 16 * 1024 * 1024))
    
    # Tempo de expiração da sessão (em segundos) - 24 horas padrão
    SESSION_TIMEOUT = int(os.environ.get('SESSION_TIMEOUT', 86400))


# =========================
# PARÂMETROS POR TIPO DE CARTÃO
# =========================

# Linha 1 -> # Conta p/ Vlr Bruto {Conta 1, Conta 2, Historico}
# Linha 2 -> # Conta p/ Taxa {Conta 1, Conta 2, Historico}

PARAMETROS = {
    "cielo credito": { # Nome do arquivo
        "palavra_data": "Data da venda",
        "palavra_valor": "Valor bruto",
        "palavra_taxa": "Taxa/tarifa",
        "ocorrencia": 2,
        "linha1": "4920;4918;359",
        "linha2": "1228;4920;363",
    },
    "cielo debito": { # Nome do arquivo
        "palavra_data": "Data da venda",
        "palavra_valor": "Valor bruto",
        "palavra_taxa": "Taxa/tarifa",
        "ocorrencia": 2,
        "linha1": "4922;4918;360",
        "linha2": "1228;4922;364",
    },
    "get net credito": { # Nome do arquivo
        "palavra_data": "DATA/HORA DA VENDA",
        "palavra_valor": "VALOR BRUTO",
        "palavra_taxa": "VALOR TAXA",
        "ocorrencia": 1,
        "linha1": "1191;4918;448",
        "linha2": "1228;1191;450",
    },
    "get net debito": { # Nome do arquivo
        "palavra_data": "DATA/HORA DA VENDA",
        "palavra_valor": "VALOR BRUTO",
        "palavra_taxa": "VALOR TAXA",
        "ocorrencia": 1,
        "linha1": "1192;4918;447",
        "linha2": "1228;1192;449",
    },
    "redecard credito": { # Nome do arquivo
        "palavra_data": "data da venda",
        "palavra_valor": "valor da venda atualizado",
        "palavra_taxa": "valor MDR",
        "ocorrencia": 1,
        "linha1": "4921;4918;361",
        "linha2": "1228;4921;365",
    },
    "redecard debito": { # Nome do arquivo
        "palavra_data": "data da venda",
        "palavra_valor": "valor da venda atualizado",
        "palavra_taxa": "valor MDR",
        "ocorrencia": 1,
        "linha1": "4923;4918;362",
        "linha2": "1228;4923;366",
    },
    "stone debito": { # Nome do arquivo
        "palavra_data": "DATA DA VENDA",
        "palavra_valor": "VALOR BRUTO",
        "palavra_taxa": "DESCONTO DE MDR",
        "ocorrencia": 1,
        "linha1": "5032;1270;372",
        "linha2": "1226;1270;456",
    },    
    "stone credito": { # Nome do arquivo
        "palavra_data": "DATA DA VENDA",
        "palavra_valor": "VALOR BRUTO",
        "palavra_taxa": "DESCONTO DE MDR",
        "ocorrencia": 1,
        "linha1": "5032;1271;372",
        "linha2": "1226;1271;456",
    },
    "gmad get net": { # Nome do arquivo
        "palavra_data": "Data/Hora \nda Venda",
        "palavra_valor": "Valor Bruto",
        "palavra_taxa": "Valor da Taxa \ne/ou Tarifa",
        "ocorrencia": 1,
        "linha1": "1191;142;448",
        "linha2": "1228;1191;450",
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo

    },
    "truckpag credito": { # Nome do arquivo
        "palavra_data": "Data Transacao",
        "palavra_valor": "Valor Total",
        "palavra_taxa": "Taxa",
        "ocorrencia": 1,
        "linha1": "5020;4918;354",
        "linha2": "1228;5020;355",
        "complemento": "TRUCKPAG",
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo


    },
    "ticket credito": { # Nome do arquivo
        "palavra_data": "Data / Hora",
        "palavra_valor": "Valor Bruto",
        "palavra_taxa": "Valor Descontado",
        "ocorrencia": 1,
        "linha1": "808;4918;405",
        "linha2": "1228;808;407", 
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo

    },
    "sodexo credito": { # Nome do arquivo
        "palavra_data": "Data da transação",
        "palavra_valor": "Valor bruto",
        "palavra_taxa": "Taxa",
        "ocorrencia": 1,
        "linha1": "819;4918;403", 
        "linha2": "1228;819;404",  
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo

    },
    "shellbox credito": { # Nome do arquivo
        "palavra_data": "Data da Transação",
        "palavra_valor": "Valor do Pagamento",
        "palavra_taxa": "Taxa",
        "ocorrencia": 1,
        "linha1": "5013;4918;354", 
        "linha2": "1228;5013;355",  
        "complemento": "SHELL",
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo

    },
    "safra credito": { # Nome do arquivo
        "palavra_data": "Data da Venda",
        "palavra_valor": "Valor Bruto da Venda",
        "palavra_taxa": "Taxa",
        "ocorrencia": 1,
        "linha1": "1265;4918;493", 
        "linha2": "1228;1265;355",  
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo
    },
    "safra debito": { # Nome do arquivo
        "palavra_data": "Data da Venda",
        "palavra_valor": "Valor Bruto da Venda",
        "palavra_taxa": "Taxa",
        "ocorrencia": 1,
        "linha1": "1266;4918;492", 
        "linha2": "1228;1266;355",  
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo
    },
    "pagseguro": { # Nome do arquivo
        "palavra_data": "Data da Transação",
        "palavra_valor": "Valor Bruto",
        "palavra_taxa": "Valor Taxa",
        "ocorrencia": 1,
        "linha1": "1238;4918;354", 
        "linha2": "1228;1238;355",  
        "complemento": "PAG SEGURO",
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo
    },
    "alelo credito": { # Nome do arquivo
        "palavra_data": "Data da Venda",
        "palavra_valor": "Valor Bruto",
        "palavra_taxa": "Taxa",
        "ocorrencia": 1,
        "linha1": "4920;4918;359", 
        "linha2": "1228;4920;363",  
        "ignorar_valores_negativos": True,  # Ignora linhas com valor bruto negativo
    },
    "mc agro sicredi debito e credito": {
        "palavra_data": "Data da venda",
        "palavra_valor": "Valor bruto",
        "palavra_taxa": "Valor da taxa",
        "ocorrencia": 1,
        # Configuração para identificar débito/crédito por coluna
        "identificar_por_coluna": {
            "nome_coluna": "Produto",  # Nome da coluna que tem "Debito" ou "Credito"
            "debito": {
                # Valores possíveis que indicam débito (case-insensitive)
                "valores": ["Débito"],
                "linha1": "5017;4918;486",
                "linha2": "1228;5017;487",
            },
            "credito": {
                # Valores possíveis que indicam crédito (case-insensitive)
                "valores": ["Crédito à vista", "Parcelado Lojista"],
                "linha1": "5016;4918;485",
                "linha2": "1228;5016;488",
            }
        }
    },
    "somavet sicredi debito e credito": {
        "palavra_data": "Data da venda",
        "palavra_valor": "Valor bruto",
        "palavra_taxa": "Valor da taxa",
        "ocorrencia": 1,
        # Configuração para identificar débito/crédito por coluna
        "identificar_por_coluna": {
            "nome_coluna": "Produto",  # Nome da coluna que tem "Debito" ou "Credito"
            "debito": {
                # Valores possíveis que indicam débito (case-insensitive)
                "valores": ["Débito"],
                "linha1": "428;4918;483",
                "linha2": "1228;428;481",
            },
            "credito": {
                # Valores possíveis que indicam crédito (case-insensitive)
                "valores": ["Crédito à vista", "Parcelado Lojista"],
                "linha1": "429;4918;482",
                "linha2": "1228;429;480",
            }
        }
    },
    "instituto emilia cielo": {
        "palavra_data": "Data da venda",
        "palavra_valor": "Valor bruto",
        "palavra_taxa": "Taxa/tarifa",
        "ocorrencia": 2,
        # Configuração para identificar débito/crédito por coluna
        "identificar_por_coluna": {
            "nome_coluna": "Forma de pagamento",  # Nome da coluna que tem "Debito" ou "Credito"
            "debito": {
                # Valores possíveis que indicam débito (case-insensitive)
                "valores": ["Débito à vista", "Débito pré-pago"],
                "linha1": "4922;4918;360",
                "linha2": "1228;4922;364",
            },
            "credito": {
                # Valores possíveis que indicam crédito (case-insensitive)
                "valores": ["Crédito à vista", "Crédito parcelado loja"],
                "linha1": "4920;4918;359",
                "linha2": "1228;4920;363",
            }
        }
    },
    "clelia rocha sicred": {
        "palavra_data": "Data",
        "palavra_valor": "Valor total",
        "palavra_taxa": "Taxa",
        "ocorrencia": 2,
        # Configuração para identificar débito/crédito por coluna
        "identificar_por_coluna": {
            "nome_coluna": "Tipo de venda",  # Nome da coluna que tem "Debito" ou "Credito"
            "debito": {
                # Valores possíveis que indicam débito (case-insensitive)
                "valores": ["DEBIT"],
                "linha1": "5009;4918;354",
                "linha2": "1228;5009;355",
            },
            "credito": {
                # Valores possíveis que indicam crédito (case-insensitive)
                "valores": ["CREDIT"],
                "linha1": "5010;4918;354",
                "linha2": "1228;5010;355",
            }
        }
    },
    "sacramento stone": {
        "palavra_data": "DATA DA VENDA",
        "palavra_valor": "VALOR BRUTO",
        "palavra_taxa": "DESCONTO DE MDR",
        "ocorrencia": 2,
        # Configuração para identificar débito/crédito por coluna
        "identificar_por_coluna": {
            "nome_coluna": "PRODUTO",  # Nome da coluna que tem "Debito" ou "Credito"
            "debito": {
                # Valores possíveis que indicam débito (case-insensitive)
                "valores": ["Débito", "Débito Pré-pago"],
                "linha1": "1270;4918;455",
                "linha2": "1228;1270;456",
            },
            "credito": {
                # Valores possíveis que indicam crédito (case-insensitive)
                "valores": ["Crédito", "Crédito Pré-pago"],
                "linha1": "1271;4918;455",
                "linha2": "1228;1271;456",
            }
        }
    },
    "unicred debito e credito": {
        "palavra_data": "Data Trans.",
        "palavra_valor": "Valor Bruto",
        "palavra_taxa": "Desconto",
        "ocorrencia": 1,
        # Configuração para identificar débito/crédito por coluna
        "identificar_por_coluna": {
            "nome_coluna": "Tipo",  # Nome da coluna que tem "Debito" ou "Credito"
            "debito": {
                # Valores possíveis que indicam débito (case-insensitive)
                "valores": ["DÉBITO"],
                "linha1": "5011;4918;350",
                "linha2": "1227;5011;350",
            },
            "credito": {
                # Valores possíveis que indicam crédito (case-insensitive)
                "valores": ["CRÉDITO"],
                "linha1": "5012;4918;350",
                "linha2": "1227;5012;350",
            }
        },
        "complemento_debito": "VENDA CARTÃO UNICRED DEBITO",
        "complemento_debito_desconto": "COMISSÃO CARTÃO UNICRED DEBITO",
        "complemento_credito": "VENDA CARTÃO UNICRED CREDITO",
        "complemento_credito_desconto": "COMISSÃO CARTÃO UNICRED CREDITO"
    },
}

# =========================
# MENSAGENS PERSONALIZADAS
# =========================
MESSAGES = {
    'upload_success': '{count} arquivo(s) enviado(s) com sucesso',
    'upload_error': 'Nenhum arquivo enviado',
    'processing_complete': 'Processamento concluído',
    'processing_error': 'Erro ao processar arquivos',
    'file_not_found': 'Arquivo não encontrado',
    'file_not_processed': 'Arquivo ainda não foi processado',
    'clear_success': 'Arquivos limpos com sucesso',
    'type_not_identified': 'Tipo de cartão não identificado no nome do arquivo. O nome deve conter: "cielo credito", "cielo debito", "get net credito", "get net debito", "redecard credito", "redecard debito" ou "stone"',
    'column_not_found': '✗ Coluna {column} não encontrada (procurava por: "{word}")',
    'type_column_not_found': '⚠️ Coluna "{column}" não encontrada na planilha. Esta coluna é necessária para identificar se cada transação é débito ou crédito. Verifique se o nome da coluna está correto no arquivo config.py',
}
