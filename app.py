from flask import Flask, render_template, request, redirect
import pandas as pd

app = Flask(__name__)

PLANILHA = "Media_Alunos.escola.xlsx"

def ler_alunos():
    df = pd.read_excel(PLANILHA)
    alunos = []

    for i, linha in df.iterrows():

        notas = []
        for numero in range(1,7):          
            coluna = "Nota" + str(numero)
            if coluna in df.columns:
                notas.append(float(linha[coluna]))

        aluno = {
            "indice"     : i,
            "nome"       : linha["Nome"],
            "serie"      : int(linha["serie"]) if pd.notna(linha["serie"]) else 0,    
            "media"      : float(linha["Media "]),
            "frequencia" : float(linha["Frequencia"]),
            "ocorrencias": int(linha["Ocorrencias"]),
            "baixa_renda": bool(linha["Baixa_renda"]),
            "trabalha"   : bool(linha["Trabalha"]),
            "notas"      : notas
        }
        alunos.append(aluno)

    return alunos

def classificar_risco(media, frequencia, ocorrencias, baixa_renda, trabalha):

    if frequencia < 60:
        return "VERMELHO"

    if media < 5 and (baixa_renda or trabalha):
        return "VERMELHO"

    if ocorrencias >= 3 and frequencia < 75:
        return "VERMELHO"

    if frequencia < 75:
        return "AMARELO"

    if media < 6:
        return "AMARELO"

    if ocorrencias >= 1 and (baixa_renda or trabalha):
        return "AMARELO"

    if baixa_renda and trabalha:
        return "AMARELO"

    return "VERDE"


def analisar_notas(notas):

    soma    = 0
    validas = 0

    for nota in notas:
        if nota == 0:
            validas = validas + 1
            continue        
        soma    = soma + nota
        validas = validas + 1

    if validas > 0:
        media_real = round(soma / validas, 1)
    else:
        media_real = 0

    # A cada nota válida calcula a média até aquele momento
    medias_progressivas = []
    soma_ate            = 0
    count_ate           = 0

    for nota in notas:
        if nota == 0:
            continue

        soma_ate  = soma_ate + nota
        count_ate = count_ate + 1
        media_ate = round(soma_ate / count_ate, 1)
        medias_progressivas.append(media_ate)

    
    notas_validas   = []
    quedas_seguidas = 0
    alerta_critico  = False

    for nota in notas:
        if nota != 0:
            notas_validas.append(nota)

    for i in range(1, len(notas_validas)):  
        if notas_validas[i] < notas_validas[i - 1]:
            quedas_seguidas = quedas_seguidas + 1

            if quedas_seguidas >= 3:
                alerta_critico = True
                break           
        else:
            quedas_seguidas = 0 

    if len(notas_validas) < 2:
        tendencia = "ESTÁVEL"
    else:
        subidas = 0
        quedas  = 0

        for i in range(1, len(notas_validas)):
            if notas_validas[i] > notas_validas[i - 1]:
                subidas = subidas + 1
            elif notas_validas[i] < notas_validas[i - 1]:
                quedas = quedas + 1

        if subidas > quedas:
            tendencia = "MELHORANDO"
        elif quedas > subidas:
            tendencia = "PIORANDO"
        else:
            tendencia = "ESTÁVEL"

    return {
        "media_real"          : media_real,
        "medias_progressivas" : medias_progressivas,
        "tendencia"           : tendencia,
        "alerta_critico"      : alerta_critico,
        "notas_validas"       : notas_validas
    }


@app.route("/")
def index():
    alunos = ler_alunos()

    for aluno in alunos:
        aluno["risco"] = classificar_risco(
            aluno["media"],
            aluno["frequencia"],
            aluno["ocorrencias"],
            aluno["baixa_renda"],
            aluno["trabalha"]
        )
        aluno["analise"] = analisar_notas(aluno["notas"])

    return render_template("index.html", alunos=alunos)


@app.route("/aluno/<int:indice>")
def ver_aluno(indice):
    alunos = ler_alunos()
    aluno  = alunos[indice]

    aluno["risco"]   = classificar_risco(
        aluno["media"], aluno["frequencia"], aluno["ocorrencias"],
        aluno["baixa_renda"], aluno["trabalha"]
    )
    aluno["analise"] = analisar_notas(aluno["notas"])

    return render_template("aluno.html", aluno=aluno)


@app.route("/novo", methods=["GET", "POST"])
def novo_aluno():
    if request.method == "POST":
        df = pd.read_excel(PLANILHA)

        notas = []
        for numero in range(1,7):
            valor = request.form.get("nota" + str(numero), "")
            if valor == "":
                notas.append(0.0)
            else:
                notas.append(float(valor))

        nova_linha = {
            "Nome"       : request.form["nome"],
            "Media "     : round(sum(n for n in notas if n != 0) /
                                 max(sum(1 for n in notas if n != 0), 1), 1),
            "Frequencia" : float(request.form["frequencia"]),
            "Ocorrencias": int(request.form["ocorrencias"]),
            "Baixa_renda": "baixa_renda" in request.form,
            "Trabalha"   : "trabalha"    in request.form,
            "Nota1": notas[0], "Nota2": notas[1], "Nota3": notas[2],
            "Nota4": notas[3], "Nota5": notas[4], "Nota6": notas[5],
        }

        df = pd.concat([df, pd.DataFrame([nova_linha])], ignore_index=True)
        df.to_excel(PLANILHA, index=False)

        return redirect("/")

    return render_template("novo.html")

@app.route("/excluir/<int:indice>")
def excluir_aluno(indice):
    df = pd.read_excel(PLANILHA)
    df = df.drop(index=indice)
    df.to_excel(PLANILHA, index=False)
    return redirect("/")

@app.route("/relatorio")
def gerar_relatorio():
    alunos = ler_alunos()
    
    # Dicionário para guardar o relatório separado por série
    relatorios_por_serie = {}

    for aluno in alunos:
        serie = aluno["serie"]
        
        # Se a série ainda não está no dicionário, cria a estrutura para ela
        if serie not in relatorios_por_serie:
            relatorios_por_serie[serie] = {
                "total_alunos": 0,
                "risco_vermelho": 0,
                "risco_amarelo": 0,
                "risco_verde": 0,
                "tendencia_piorando": 0,
                "lista_risco_critico": [],
                "lista_piorando": []
            }

        # 1. Classifica o risco
        risco = classificar_risco(
            aluno["media"],
            aluno["frequencia"],
            aluno["ocorrencias"],
            aluno["baixa_renda"],
            aluno["trabalha"]
        )
        
        # 2. Analisa a tendência de notas
        analise = analisar_notas(aluno["notas"])
        tendencia = analise["tendencia"]
        
        # 3. Atualiza os contadores da série específica
        relatorios_por_serie[serie]["total_alunos"] += 1
        
        if risco == "VERMELHO":
            relatorios_por_serie[serie]["risco_vermelho"] += 1
            relatorios_por_serie[serie]["lista_risco_critico"].append(aluno["nome"])
        elif risco == "AMARELO":
            relatorios_por_serie[serie]["risco_amarelo"] += 1
        elif risco == "VERDE":
            relatorios_por_serie[serie]["risco_verde"] += 1
            
        # 4. Atualiza os contadores de Tendência da série
        if tendencia == "PIORANDO":
            relatorios_por_serie[serie]["tendencia_piorando"] += 1
            relatorios_por_serie[serie]["lista_piorando"].append(aluno["nome"])

    # Ordena o dicionário pelas chaves (séries) para mostrar a 8ª antes da 9ª no HTML
    relatorios_ordenados = dict(sorted(relatorios_por_serie.items()))

    return render_template("relatorio.html", relatorios=relatorios_ordenados)

if __name__ == "__main__":
    app.run(debug=True)
