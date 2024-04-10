def enviar_email():
    """Realiza o envio de um relatório após as 18:00"""
    
    # Utilizada para configurar o envio de email após as 18:00
    hora_de_corte = datetime.strptime(f"{date.today()} 18:00:00", "%Y-%m-%d %H:%M:%S") 
    hora_atual = datetime.now()

    if hora_atual > hora_de_corte:
        lista_contatos = ["a.alves@perkons.com, alexander.s@perkons.com, fernando.b@perkons.com"]
        email = Emailer(EMAIL_ADDRESS, EMAIL_PASSWORD)
        caminho_disponiblidade = os.path.join(
            diretorio_disponibilidade, "DISPONIBILIDADE 2024.xlsx"
        )
        disponibilidade = openpyxl.load_workbook(caminho_disponiblidade, data_only=True)
        locale.setlocale(locale.LC_ALL, "pt_br")
        data_de_hoje = datetime.now()
        str_data_de_hoje = data_de_hoje.strftime("%Y-%m-%d 00:00:00")
        mes_numero = data_de_hoje.month
        mes_atual = calendar.month_name[mes_numero].upper()
        sheet_mes_atual = disponibilidade[mes_atual]
        mensagem = ""

        for linha in sheet_mes_atual.iter_rows(min_row=2, values_only=True):
            data = str(linha[0])
            if data == str_data_de_hoje:
                mensagem = f"""
                 <style>
                    table {{
                    width: 100%;
                    font-size: 10px;
                    border: 1px gray solid;
                    }}
                    th,
                    td {{
                    border: 1px black solid;
                    height: 50px;
                    text-align: center;
                    }}
                    .content {{
                        padding: 5px;
                        border-radius: 10px;
                        background-color: lightgray;
                        text-align:center;
                    }}
                 </style>
                <div class="content">
                <h1 text-align="center"> Relatório de Disponibilidade DETRAN-ES <h1>
                <table>
                    <thead>
                        <th> Amostra 1</th>
                        <th>Amostra 2</th>
                        <th>Amostra 3</th>
                        <th> Amostra 4</th>
                        <th>Amostra 5</th>
                        <th>Amostra 6</th>
                        <th>Média do Dia</th>
                    </thead>
                    <tbody>
                        <tr>
                            <td>{linha[1]} / {linha[2]} / {verifica_se_vazio(linha[3])}</td>
                            <td>{linha[4]} / {linha[5]} / {verifica_se_vazio(linha[6])}</td>
                            <td>{linha[7]} / {linha[8]} / {verifica_se_vazio(linha[9])}</td>
                            <td>{linha[10]} / {linha[11]} / {verifica_se_vazio(linha[12])}</td>
                            <td>{linha[13]} / {linha[14]} / {verifica_se_vazio(linha[15])}</td>
                            <td>{linha[16]} / {linha[17]} / {verifica_se_vazio(linha[18])}</td>
                            <td>{verifica_se_vazio(linha[19])}</td>                    
                        </tr>
                    </tbody>
                    </table>
                    </div>
                """
        email.definir_conteudo(topico=f"Relatório de Disponibilidade {date.today().strftime("%d/%m/%y")} ",
                           email_remetente="andre@andrealves.eng.br",
                           lista_contatos=lista_contatos,
                           conteudo_email=mensagem)
        enviar_mensagem_whatsapp(linha[19])
        try:
            email.enviar_email(intervalo_em_segundos=5)
        except Exception as e:
            print(e.args)
    else:
        # envio de mensagem para motirar se foi realizada corretamente.
        email = Emailer(EMAIL_ADDRESS, EMAIL_PASSWORD)
        lista_contatos = ["a.alves@perkons.com"]
        email.definir_conteudo(topico=f"Relatório de Disponibilidade executado com sucesso ",
                           email_remetente="andre@andrealves.eng.br",
                           lista_contatos=lista_contatos,
                           conteudo_email="Executado com sucesso !")
        
        try:
            email.enviar_email(intervalo_em_segundos=5)
            enviar_mensagem_whatsapp("Automação de Disponibilidade executada com sucesso !")
        except Exception as e:
            print(e.args)
