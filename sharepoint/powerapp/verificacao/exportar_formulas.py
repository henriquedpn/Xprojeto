# Exporta as fórmulas geradas (com o escopo de galeria) para o verificador de tipos.
# Uso: python exportar_formulas.py <pasta do gerar_powerapp.py> <saida.json>
import json, sys
sys.path.insert(0, sys.argv[1]); import gerar_powerapp as g
lista, du = g.ler_definicao(g.DEFINICAO); f = g.nome_fonte(lista)
saida = []
def walk(c, escopo, tela):
    for k, v in c.props.items():
        saida.append({"controle": c.nome, "modelo": c.modelo, "prop": k, "formula": v, "escopo": (c.nome if (c.modelo.startswith("gallery") and k == "OnSelect") else escopo), "tela": tela})
    for ch in c.filhos:
        walk(ch, c.nome if c.modelo.startswith("gallery") else escopo, tela)
for t in [g.app_onstart(), g.tela_projetos(f), g.tela_formulario(f, du), g.tela_dashboard(f)]:
    walk(t, None, t.nome)
json.dump(saida, open(sys.argv[2], "w"), ensure_ascii=False)
print(len(saida), "fórmulas exportadas")
