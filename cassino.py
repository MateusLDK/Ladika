import random
import time
import sys
import os

def limpar_terminal():
    """Limpa o terminal de forma compatível com Windows e Unix"""
    os.system('cls' if os.name == 'nt' else 'clear')

def formatar_emoji(emoji):
    """Formata cada emoji para ocupar 3 espaços fixos"""
    return f" {emoji} "

def caça_níquel_continuo():
    # Lista de emojis para o caça-níquel
    emojis = ["💰", "💎", "🍒", "🎰", "🔔", "🍀", "🏆", "♦️", "⭐"]

    while True:
        print("\n🎰 CAÇA-NÍQUEL PYTHON 🎰")
        print("Pressione Enter para girar ou 'q' + Enter para sair...")
        entrada = input()
        limpar_terminal()
        if entrada.lower() == 'q':
            print("\nObrigado por jogar! Até a próxima! 🎰\n")
            break
        
        # Animação dos emojis girando
        #print("\nGirando...\n")
        
        start_time = time.time()
        while time.time() - start_time < 1:  # Rola por 1 segundo
            # Gera 3 emojis aleatórios
            resultado = [random.choice(emojis) for _ in range(3)]
            # Escreve na mesma linha (\r) e força a saída imediata (flush)
            sys.stdout.write("\r" + "|".join([formatar_emoji(e) for e in resultado]))
            sys.stdout.flush()
            time.sleep(0.1)  # Pequeno delay entre as rotações
        
        # Resultado final (depois de 1 segundo)
        resultado_final = [random.choice(emojis) for _ in range(3)]
        print("\r" + "|".join([formatar_emoji(e) for e in resultado_final]) + "  ← RESULTADO FINAL!")
        
        # Verifica se todos são iguais (jackpot)
        if len(set(resultado_final)) == 1:
            print("\n🎉 JACKPOT! TODOS IGUAIS! 🎉")
        # Verifica se há pelo menos 2 iguais
        elif len(set(resultado_final)) == 2:
            print("\n👍 Quase! Dois iguais! 👍")
        

# Executa o caça-níquel contínuo
caça_níquel_continuo()