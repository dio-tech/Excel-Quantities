class Info:
    def __init__(self):
        self.choice = None
        self.error = "Escolher um número válido!"
    
    def get_choice(self):
        while self.choice == None:
            print('1 - Adicionar informação ao monday\n2 - Editar manualmente ficheiro excel')
            choice = input(':')
            try:
                self.choice = int(choice)
            except:
                print(self.error)
        return self.choice