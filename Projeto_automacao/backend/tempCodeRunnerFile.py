elementos = self.driver.find_elements(By.CLASS_NAME, "tela-vazia")

        if elementos and elementos[0].is_displayed():
            self.itens_verificados.append(self.leitura.lista_pecas[peca].copy())
        else:
            pass    