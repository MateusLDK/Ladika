import random
from datetime import datetime

caixas = {}

mensagens_nonsense = [
    "[SISTEMA]: A caixa parece mais leve. Alguém levou algo?",
    "[SISTEMA]: Uma nova dúvida apareceu sozinha.",
    "[SISTEMA]: A Caixa se perguntou se era mesmo necessária.",
    "[SISTEMA]: Um segredo escapou por entre as frestas.",
    "[SISTEMA]: Caixa duplicada no plano imaginário detectada.",
    "[SISTEMA]: Os itens da caixa estão cochichando entre si.",
    "[SISTEMA]: A caixa desenvolveu consciência espacial e sabe que está sendo observada.",
    "[SISTEMA]: Um paradoxo foi detectado: a caixa contém a si mesma em uma de suas dimensões.",
    "[SISTEMA]: A tampa da caixa piscou para você. Foi sua imaginação?",
    "[SISTEMA]: A caixa está emitindo um cheiro de nostalgia de infância.",
    "[SISTEMA]: Os itens se reorganizaram sozinhos durante a noite.",
    "[SISTEMA]: Uma caixa paralela tentou se comunicar através de padrões de sombras.",
    "[SISTEMA]: A caixa está mais pesada hoje. Talvez acumule poeira de estrelas.",
    "[SISTEMA]: Um item fugiu para se tornar poeta em outra dimensão.",
    "[SISTEMA]: A caixa está emitindo um zumbido na frequência do pensamento criativo.",
    "[SISTEMA]: Detectada atividade onírica nos itens armazenados - alguns podem estar sonhando.",
    "[SISTEMA]: A caixa está fazendo perguntas retóricas para o vazio."
]

def criar_caixa(nome):
    if nome in caixas:
        print(f"⚠️ Caixa '{nome}' já existe.")
    else:
        caixas[nome] = []
        print(f"📦 Criada caixa '{nome}'.")

def guardar(item, nome):
    if nome not in caixas:
        print(f"❌ Caixa '{nome}' não existe.")
        return
    caixas[nome].append(item)
    print(f"📥 Guardado '{item}' em '{nome}'.")

def abrir_caixa(nome):
    if nome not in caixas:
        print(f"❌ Caixa '{nome}' não existe.")
        return
    print(f"📂 Conteúdo da caixa '{nome}':")
    if caixas[nome]:
        for i, item in enumerate(caixas[nome], 1):
            print(f"  {i}. {item}")
    else:
        print("  (vazia)")

def remover_item(nome, indice):
    if nome not in caixas:
        print(f"❌ Caixa '{nome}' não existe.")
        return
    try:
        item = caixas[nome].pop(indice - 1)
        print(f"🗑️ Removido '{item}' de '{nome}'.")
    except IndexError:
        print("⚠️ Índice inválido.")

def listar_caixas():
    if not caixas:
        print("🚫 Nenhuma caixa criada ainda.")
        return
    print("📦 Caixas existentes:")
    for nome in caixas:
        print(f"  - {nome} ({len(caixas[nome])} itens)")

def comando_maluco():
    if random.random() < 0.3:
        print(random.choice(mensagens_nonsense))

    if random.random() < 0.1:  # 30% de chance
        comentarios_caóticos()


def auto_preencher(num_caixas=3, max_itens=5):
    temas_caixas = [
    # Temas cósmicos/abstratos
    "Cosmogonias", "Eclipses", "Nebulosas", "Quasares", "Universos",
    "Singularidades", "Galáxias", "Órbitas", "Cometas", "Astros",
    
    # Temas de memória/tempo
    "Arquivos", "Crônicas", "Anais", "Relíquias", "Fósseis",
    "Anacronismos", "Profecias", "Ucronias", "Distopias", "Futuros",
    
    # Temas poéticos/etéreos
    "Epifanias", "Miragens", "Alquimias", "Metamorfoses", "Transfigurações",
    "Êxtases", "Delírios", "Fantasmagorias", "Aparições", "Iluminações",
    
    # Temas cotidianos distorcidos
    "Sussurros", "Rastros", "Reflexos", "Sombras", "Ecos",
    "Lembranças", "Cheiros", "Pulsos", "Sinais", "Vestígios",
    
    # Temas conceituais
    "Paradoxos", "Enigmas", "Arcanos", "Hieróglifos", "Manuscritos",
    "Cifras", "Algoritmos", "Equações", "Fórmulas", "Teoremas"
]
    
    adjetivos = [
        "Abstratos", "Alquímicos", "Anacrônicos", "Aurorais", "Brilhantes",
        "Caleidoscópicos", "Cambiantes", "Cintilantes", "Cósmicos", "Cristalinos",
        "Desbotados", "Dissonantes", "Efímeros", "Enigmáticos", "Etéreos",
        "Fantasiosos", "Fugidios", "Fosforescentes", "Galácticos", "Híbridos",
        "Inefáveis", "Invisíveis", "Irradiados", "Lânguidos", "Lúdicos",
        "Luminescentes", "Misteriosos", "Nebulosos", "Oníricos", "Ocultos",
        "Pálidos", "Pulsantes", "Quiméricos", "Ressonantes", "Sombrios",
        "Sussurrantes", "Tremeluzentes", "Translúcidos", "Utópicos", "Vagantes",
        "Vibrantes", "Viscosos", "Vazio", "Xamânicos", "Zumbis",
        "Âmbar", "Ébrios", "Ígneos", "Órfãos", "Únicos"
    ]
    
    categorias_itens = {
    # Objetos enigmáticos
    "artefatos": [
        "Chave de cristal que abre portas invisíveis",
        "Pena que escreve sozinha quando lua está cheia",
        "Caixa dentro da caixa (contém o próprio universo)",
        "Relógio que marca o tempo de sonhos",
        "Espelho que reflete memórias alheias",
        "Bússola que aponta para desejos não confessados"
    ],
    
    # Pensamentos vivos
    "ideias": [
        "Pensamento que muda de forma quando observado",
        "Dúvida que pesa 327 gramas",
        "Teorema que prova a existência de fadas",
        "Eureka! (embalado a vácuo)",
        "Argumento circular (cuidado para não cair dentro)",
        "Hipótese que brilha no escuro"
    ],
    
    # Fragmentos sensoriais
    "sensações": [
        "Cheiro de chuva de um planeta desconhecido",
        "Tato de nuvem armazenado em frasco",
        "Gosto de lágrima de alegria cristalizada",
        "Sussurro preservado em âmbar",
        "Arrepio que se repete em loop",
        "Vislumbre de déjà vu compactado"
    ],
    
    # Mistérios temporais
    "anacronismos": [
        "Carta datada de 32 de fevereiro de 2157",
        "Fotografia de um evento que nunca aconteceu",
        "Ingresso para o fim dos tempos (já usado)",
        "Mapa de uma cidade que ainda será fundada",
        "Receita culinária escrita em língua extinta",
        "Semente de árvore que ainda será evolúida"
    ],
    
    # Resíduos cósmicos
    "cosmológicos": [
        "Pó de estrela cadente (manuseio cuidadoso)",
        "Fragmento de buraco negro (não expor à luz)",
        "Amostra de silêncio interestelar",
        "Pedacinho da borda do universo",
        "Gotícula de tempo espacial",
        "Filete de energia escura em bisnaga"
    ],
    
    # Objetos cotidianos transformados
    "distorcidos": [
        "Garfo que só corta sombras",
        "Xícara que sempre está cheia de lembranças",
        "Meia que desaparece paulatinamente",
        "Lâmpada que ilumina pensamentos",
        "Guarda-chuva que protege de más ideias",
        "Livro cujas páginas são feitas de tempo"
    ],
    
    # Emoções tangíveis
    "sentimentos": [
        "Pacote de saudade concentrada",
        "Raiva engarrafada (não agitar)",
        "Alegria em pó (só adicionar água)",
        "Medo dobrado cuidadosamente",
        "Vontade de viajar (embalada a vácuo)",
        "Pedaço de coração partido (para remontagem)"
    ],
    
    # Mistérios linguísticos
    "linguísticos": [
        "Palavra que não pode ser pronunciada",
        "Pergunta cuja resposta queima a língua",
        "Poema escrito em código genético",
        "Silêncio entre parênteses",
        "Grito preservado em formol",
        "Diálogo entre átomos (traduzido)"
    ],
    
    # Seres imaginários
    "entidades": [
        "Asa de anjo de guarda (esquerda)",
        "Risada de duende (em conserva)",
        "Pegada de unicórnio não identificado",
        "Sombra de um ser bidimensional",
        "Nome verdadeiro de um fantasma",
        "Último suspiro de um dragão (framboesa)"
    ],
    
    # Natureza surreal
    "naturais": [
        "Raio de sol engarrafado em inverno",
        "Flor que desabrocha só em sonhos",
        "Pedaço de horizonte (dobrável)",
        "Amanhecer em pó (só adicionar água)",
        "Maré congelada em cubos",
        "Eclipse lunar em sachê"
    ],
    
    # Tecnologias impossíveis
    "invenções": [
        "Máquina de desfazer lembranças",
        "Aparador de paradoxos (pilhas não inclusas)",
        "Gerador de finais felizes alternativos",
        "Traduutor de línguas inexistentes",
        "Telescópio que vê através do tempo",
        "Tinta que apaga a realidade"
    ],
    
    # Relíquias pessoais
    "pessoais": [
        "Primeiro amor (desidratado)",
        "Último adeus (não perecível)",
        "Beijo não dado (embalado a vácuo)",
        "Momento exato em que tudo mudou",
        "Pedido de desculpas engasgado",
        "Segredo que nunca foi contado"
    ]
}
    
    for _ in range(num_caixas):
        # Gerar nome criativo para caixa
        tema = random.choice(temas_caixas)
        adj = random.choice(adjetivos)
        nome_caixa = f"{tema} {adj}"
        
        # Criar a caixa
        criar_caixa(nome_caixa)
        
        # Adicionar itens aleatórios
        num_itens = random.randint(1, max_itens)
        for _ in range(num_itens):
            categoria = random.choice(list(categorias_itens.keys()))
            item = random.choice(categorias_itens[categoria])
            
            # 30% de chance de adicionar um timestamp
            if random.random() < 0.3:
                hora = datetime.now().strftime("%H:%M")
                item = f"{item} [{hora}]"
            
            # 20% de chance de adicionar um modificador
            if random.random() < 0.2:
                modificadores = ["leve", "pesado", "vibrante", "desbotado", "quente", "frio"]
                item = f"{item} ({random.choice(modificadores)})"
            
            guardar(item, nome_caixa)
        
        # Mensagem especial
        print(f"✨ Caixa '{nome_caixa}' auto-preenchida com {num_itens} itens misteriosos!")
    
    # Efeito especial aleatório
    if random.random() < 0.4:
        print("[SISTEMA]: Alguns itens parecem ter se movido sozinhos durante o processo...")


def comentarios_caóticos():
    if not caixas:
        return
    
    # Escolhe uma caixa aleatória que tenha itens
    caixas_com_itens = [nome for nome, itens in caixas.items() if len(itens) >= 2]
    if not caixas_com_itens:
        return
    
    nome_caixa = random.choice(caixas_com_itens)
    itens = caixas[nome_caixa]
    item1, item2 = random.sample(itens, 2)
    
    # Lista de frases caóticas
    interacoes = [
        f"[SISTEMA]: |{item1}| e |{item2}| estão tendo uma discussão filosófica sobre a natureza das caixas...",
        f"[SISTEMA]: Algo me diz que |{item1}| não combina nem um pouco com |{item2}| (mas quem sou eu pra julgar?)",
        f"[SISTEMA]: |{item1}| e |{item2}| formaram uma aliança secreta contra os outros itens da caixa!",
        f"[SISTEMA]: Cuidado! |{item1}| e |{item2}| estão prestes a criar um paradoxo ontológico!",
        f"[SISTEMA]: |{item1.split('[')[0]}| e |{item2.split('[')[0]}| saíram no tapa por causa de um enroladinho de vento",
        f"[SISTEMA]: Estranho... |{item1}| parece estar evitando |{item2}| desde o incidente",
        f"[SISTEMA]: {random.choice(['Essa combinação', '|' + item1 + '| com |' + item2 + '|'])} não parece ser uma boa ideia",
        f"[SISTEMA]: Se você ouvir sussurros, é só |{item1}| contando segredos para |{item2}|",
        f"[SISTEMA]: |{item1}| e |{item2}| estão competindo pelo título de 'Item Mais Estranho da Caixa'",
        f"[SISTEMA]: Eu avisei que não deviam colocar |{item1}| junto com |{item2}|... agora estão cantando em dueto!",
        f"[SISTEMA]: |{item1}| está dando aula de alquimia para |{item2}| (os resultados são imprevisíveis)",
        f"[SISTEMA]: Alerta! |{item1}| e |{item2}| estão tentando criar uma startup juntos!"
    ]
    
    # 20% de chance de ter um comentário extra absurdo
    if random.random() < 0.2:
        interacoes.extend([
            f"[SISTEMA EM PÂNICO]: |{item1}| COMEU |{item2}|!!! ... Espera, não, estão ali... ou não?",
            f"[SISTEMA]: |{item1}| e |{item2}| sumiram! Brincadeira... ou será que não? *sons de risadinhas*",
            f"[SISTEMA CRÍPTICO]: O terceiro ato entre |{item1[:15]}|... e |{item2[:15]}|... foi previsto nos antigos pergaminhos..."
        ])
    
    print(random.choice(interacoes))
    # 10% de chance do comentário afetar os itens
    if random.random() < 0.1:
        if random.choice([True, False]):
            itens.remove(item1)
            print(f"⚡ |{item1}| fugiu da caixa após o comentário!")
        else:
            itens.append(f"Fusão de -{item1[:10]}-... + -{item2[:10]}-...")
            print(f"⚡ Os itens se fundiram criando algo novo e bizarro!")

def main():
    print("🧠 Organizador de Caixas Imaginárias™")
    print("Digite comandos como:")
    print("  criar Caixa_Azul")
    print("  guardar 'uma ideia brilhante' em Caixa_Azul")
    print("  abrir Caixa_Azul")
    print("  remover 1 de Caixa_Azul")
    print("  listar")
    print("  sair\n")

    while True:
        try:
            comando = input(">> ").strip()

            if comando == "sair":
                print("🫡 Arquivamento encerrado.")
                break
            elif comando.startswith("criar "):
                nome = comando[6:].strip()
                criar_caixa(nome)
            elif comando.startswith("guardar "):
                try:
                    parte1, parte2 = comando[8:].split(" em ")
                    item = parte1.strip().strip("'\"")
                    nome = parte2.strip()
                    guardar(item, nome)
                except ValueError:
                    print("⚠️ Comando inválido. Use: guardar 'item' em Caixa_X")
            elif comando.startswith("abrir "):
                nome = comando[6:].strip()
                abrir_caixa(nome)
            elif comando.startswith("remover "):
                try:
                    parte1, parte2 = comando[8:].split(" de ")
                    indice = int(parte1.strip())
                    nome = parte2.strip()
                    remover_item(nome, indice)
                except:
                    print("⚠️ Comando inválido. Use: remover 1 de Caixa_X")
            elif comando == "listar":
                listar_caixas()
            elif comando == "autopreencher":
                auto_preencher()
            elif comando == "":
                pass
            else:
                print("❓ Comando não reconhecido.")
            
            comando_maluco()

        except KeyboardInterrupt:
            print("\n🫡 Arquivamento encerrado.")
            break

if __name__ == "__main__":
    main()
