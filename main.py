from processos.Auditoria import executar_auditoria
from processos.FalsasFaltas import auditoria_sequencial
from processos.Formatacao import executar_formatacao


def main():
    print("Iniciando a auditoria de faltas...")
    executar_auditoria()

    print("Iniciando filtragem de servidores em licença")
    auditoria_sequencial()

    print("Iniciando a formatação")
    executar_formatacao()

    print("A auditoria está completa")
    print("Executar o codigo de datas no AppScript na conta do GOOGLE da CMP")
    print("https://script.google.com/home/projects/1FNDk7N6M9CQqO_p1cDBYt5js05pB_FbQqH07tmfs47qtGlKZIv4MH0kU/edit")

if __name__ == "__main__":
    main()



        