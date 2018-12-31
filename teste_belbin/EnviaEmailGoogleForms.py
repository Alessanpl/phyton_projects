# encoding: utf-8

"""
  Envia uma mensagem para cada aluno cadastrado na planilha.
  E-mails enviados pelo alessanpl@gmail.com
"""

# Modulo para manipulação de email
import smtplib
from email.mime.text import MIMEText

# Modulo pandas
import pandas as pd

# Modulo para manipulação de datas
import datetime as dt

pasta = "/Users/Alessanpl 1/PhytonFundAle/Ferramentas de Script/Enviar teste de Belbin/"

# Envia e-mail simples
def EnviaEmailTodosAlunos(assunto, de, para, aluno, mensagem):
    msg = MIMEText(mensagem)
    msg['Subject'] = assunto 
    msg['From'] = de
    msg['To'] = aluno
    msg['Cco'] = para

    # Enviando o e-mail
    Consmtp.send_message(msg)

# Logando no servidor do GMAIL
Consmtp = smtplib.SMTP("smtp.gmail.com", 587)
Consmtp.starttls()
Consmtp.login("alessanpl@gmail.com", "alessandro110")

# le planilha excel e gera um df
df = pd.read_excel(pasta + "Teste de Belbin (Responses).xlsx")

# altera os nomes das colunas do df
df.columns = ['data', 'email', 'nome', 'Ia', 'Ib', 'Ic', 'Id', 'Ie', 'If', 'Ig', 'Ih', 'IIa', 'IIb', 'IIc', 'IId', 'IIe', 'IIf', 'IIg', 'IIh', 'IIIa', 'IIIb', 'IIIc', 'IIId', 'IIIe', 'IIIf', 'IIIg', 'IIIh', 'IVa', 'IVb', 'IVc', 'IVd', 'IVe', 'IVf', 'IVg', 'IVh', 'Va', 'Vb', 'Vc', 'Vd','Ve', 'Vf', 'Vg', 'Vh','VIa', 'VIb', 'VIc', 'VId', 'VIe', 'VIf', 'VIg', 'VIh', 'VIIa', 'VIIb', 'VIIc', 'VIId', 'VIIe', 'VIIf', 'VIIg', 'VIIh']

msgcoordenador = """
    COORDENADOR
    SINERGIA E LIDERANÇA

    PAPEL: Controlar o modo com que a equipe se move em direção aos objetivos do grupo, fazendo o melhor uso dos recursos da equipe; 
    reconhecer onde se localizam as forças e fraquezas materiais e humanas da equipe e assegurar que o potencial de cada membro da equipe está sendo utilizado da melhor forma.

    CARACTERÍSTICAS:
    • Pontos Fortes: uma habilidade de comandar e inspirar entusiasmo, um senso de oportunidade e equilíbrio e a uma capacidade de comunicar-se  facilmente com os outros.
    • Pontos Fracos Toleráveis: sem capacidade intelectual ou criativa acentuada.
    • Habilidades comportamentais: clarificar, sintetizar, “trazer as pessoas para dentro”, perguntas abertas.

    DESCRIÇÃO: Há dois tipos de líderes de equipe - o “coordenador” e o “moldador”. 
    O coordenador  demonstra capacidade de controle e coordenação dos recursos entre o grupo. 
    Atua em uma base democrática e participativa, mas está pronto para assumir o controle mais diretamente, quando sente que é necessário.  
    É bom em reconhecer e utilizar os recursos entre o grupo, equilibrando suas forças e fraquezas.  Mantém a visão dos objetivos e direciona as atividades da equipe em direção a esses objetivos.

    A marca do coordenador descreve bem as suas forças, mas pode ser enganosa, como se ele não fosse necessariamente o líder escolhido da equipe. 
    Quando se encontra num cargo mais júnior, é possível que precise manter a harmonia e estrutura dentro da equipe, sem ameaçar o líder da equipe. 
    De fato, o coordenador tem muitas das qualidades dos outros membros da equipe, e precisa estar preparado para adaptar-se a necessidades específicas do grupo. 
    A despeito de sua força de ego, deverá saber em que papel atuar e quando.

    Frases e slogans que caracterizam um COORDENADOR:

    1. Vamos manter o objetivo principal em mente.
    2. Alguém mais tem alguma coisa a acrescentar a isso?
    3. Nós gostamos de chegar a um consenso antes de seguirmos adiante.
    4. Nunca assuma que silêncio significa aprovação.
    5. Eu acho que nós devemos dar uma chance a mais alguém.
    6. Delegar bem é uma arte.
    7. Gerência é a arte de conseguir outras pessoas para fazer todo o trabalho.\n\n"""

msgmoldador = """
    MOLDADOR
    CONDUZIR E DAR URGÊNCIA

    PAPEL: Molda o caminho em que os esforços da equipe são aplicados, dirigindo atenção geralmente para os objetivos e prioridades colocadas, e procura impor uma forma ou modelo nas discussões do grupo e nos resultados das atividades do grupo.

    CARACTERÍSTICAS:
    • Pontos Fortes: direção e auto-confiança.
    • Pontos fracos toleráveis: intolerância à pessoas e idéias vagas.
    • Habilidades comportamentais: desafiador, harmonizador, condutor, energizador.

    DESCRIÇÃO: Este outro tipo de líder. O moldador, está muito mais dentro do modelo tradicional de assumir a liderança de primeira. 
    Ele prefere “formatar” as decisões e atividades da equipe direta e pessoalmente.  Gosta de ação e resultados rápidos, e de pessoas que o seguem. 
    Procura o curso de ação que melhor sumarize as necessidades da situação, e tem a capacidade de sair de questões difíceis com algumas frases apropriados e decisões incisivas. 
    Ele se lançará, e as idéias que acredita serem as melhores também, com dois propósitos:

    CONCLUIR O TRABALHO SEM PERDA DE TEMPO
    SATISFAZER SUAS PRÓPRIAS NECESSIDADES DE ESTAR NO COMANDO
    O MOLDADOR não é necessariamente popular, apesar de sempre tirar o melhor da equipe; mas alcança resultados, e em muitos casos é o que é desejado.  Funciona melhor em uma equipe de pares informais.  
    Numa equipe mais formal e estruturada, a forma do coordenador é freqüentemente mais eficaz. O MOLDADOR precisa de mais auto-disciplina para ser mais coordenador. 
    Quando ele é relativamente júnior, precisa dosar sua contribuição com habilidade e diplomacia, se quiser ser bem sucedido.

    Frases e slogans que caracterizam um MOLDADOR:

    1. Apenas faça.
    2. Diga não, depois negocie.
    3. Se você diz ‘ sim, eu vou fazer’, eu espero que seja feito.
    4. Eu não estou satisfeito com nosso rendimento..
    5. Eu posso ser meio óbvio, mas pelo menos eu sou direto.
    6. Eu vou fazer as coisas andarem.
    7. Quando o caminho fica árduo, é que a gente cresce.\n\n"""

msginovador = """
    INOVADOR
    INPUT CRIATIVO

    PAPEL: Desenvolver novas idéias e estratégias com atenção especial para questões importantes. Procura possíveis aberturas  na abordagem de problemas com que o grupo possa se confrontar.

    CARACTERÍSTICAS:
    • Pontos fortes: independência de ponto de vista, muito inteligente e imaginativo.
    • Pontos fracos toleráveis: uma tendência a ser impraticável ou estar alguns momentos “nas nuvens”. Carece de boa comunicação com os outros e é defensivo quando desafiado.
    • Habilidades comportamentais: originalidade e fazer perguntas “estúpidas” (para provocar o pensamento).

    DESCRIÇÃO: O input criativo de uma equipe vem do inovador ou da pessoa de idéias. Seu ponto forte está na capacidade de desenvolver novas idéias e estratégias; uma capacidade que independe de sua habilidade profissional. 
    Ele tem uma mente fértil e ativa, uma capacidade para idéias originais, seja para novos produtos ou solução de problemas estratégicos; e quando é aceito pela equipe, pode transformar a forma da equipe pensar e contribuir imensamente para seu sucesso.

    Existe um perigo que o inovador venha rejeitar a equipe, se suas idéias forem rejeitadas, e mesmo quando está “ligado”, pode ser um colega difícil. 
    É necessário levar o inovador com cuidado (geralmente pelo coordenador), e a equipe deve ter consciência de suas forças e deve dar um apoio complementar para conseguir o melhor dele. 
    Mas sem ele, a equipe fica meio perdida, sem criatividade, com carência de brilho e raramente sobressai. De fato, os inovadores são conhecidos por exercitarem  auto- disciplina, auto-crítica, e preocupação com acordos que se fazem necessários com os colegas, e mesmo com idéias de outras pessoas.

    Frases e slogans que caracterizam um INOVADOR:

    1. Quando um problema parece não ter solução, pense sob outro ângulo.
    2. Onde existe um problema, existe uma solução.
    3. Quanto maior o problema, maior o desafio.
    4. Não perturbe, gênio trabalhando.
    5. Boas ideias sempre soam estranhas num primeiro momento.
    6. Ideias começam com um sonho.
    7. Sem inovação, não há sobrevivência.\n\n"""

msgmonitor = """
    MONITOR / AVALIADOR
    PENSAMENTO CRÍTICO

    PAPEL: Analisa problemas e avalia idéias e sugestões de forma que a equipe esteja melhor preparada para tomar decisões ajustadas.

    CARACTERÍSTICAS:
    • Pontos Fortes: habilidade de pensar criticamente, incluindo a habilidade de enxergar as   propostas confusas; uma mente objetiva.
    • Pontos Fracos Toleráveis: super crítico, pouco entusiasmado, um pouco sério demais.
    • Habilidades Comportamentais: objetivamente disciplinado, brilhante sintetizador, boa capacidade crítica [realista - precisa].

    DESCRIÇÃO: O monitor é um complemento imediato para o inovador. Seu ponto forte está na habilidade de pensar criticamente: analisar ideias e sugestões da equipe, e avaliar se podem ou não ser executadas e valor prático das mesmas, em termos do objetivo da equipe. 
    Ele traz não somente uma mente aguçada para o funcionamento da equipe, mas também uma percepção contundente, cautelosa e objetiva. 
    Ele é um estrategista e um perito, mas precisa de ideias, conhecimento do que está acontecendo no mundo de fora; precisa de perspectiva global e bom senso, que outros papéis da equipe podem prover se ele tiver que usar sua aptidão crítica com um bom propósito.

    Por natureza, tende a ser crítico demais e a diminuir a moral da equipe, e a alienar o inovador. 
    Se bem aproveitado, contudo, sua habilidade trará uma visão crítica para projetos imprudentes, ou para equipes que se empolgam demais com idéias impraticáveis. 
    Ele pode ajudar os colegas a atingirem uma ótima decisão, proporcionando informações conflitantes para os gerentes. Ele tende a ser intelectualmente competitivo com os colegas e corre o risco de domina-los em demasia.

    Frases e slogans que caracterizam um MONITOR / AVALIADOR:

    1. Eu vou pensar sobre o assunto e te dou uma resposta amanhã.
    2. Nós esgotamos todas as possibilidades ?
    3. Se não segue uma lógica, não vale a pena fazer.
    4. É melhor tomar uma decisão devagar mas acertada, do que rápida mas errada.
    5. Colocando na balança, esta parece ser a melhor opção.
    6. Vamos considerar todas as alternativas.
    7. Decisões não devem se basear somente no entusiasmo.\n\n"""

msgimplementador = """
    IMPLEMENTADOR
    ALCANÇAR A EXECUÇÃO DO TRABALHO

    PAPEL: Transformar conceitos e planos em procedimentos práticos de trabalho e por em execução sistemática e eficazmente planos acordados.

    CARACTERÍSTICAS:
    • Pontos Fortes: auto-controle, auto-disciplina, aliado a um realismo e senso prático comum.
    • Pontos Fracos Toleráveis: carência de flexibilidade e não compreensivo com novas idéias que ainda não foram submetidas a prova.
    • Habilidades Comportamentais: estruturar tarefas, clarificar metas, objetivos e táticas.

    DESCRIÇÃO: As formas de papéis do coordenador, moldador, inovador e monitor/avaliador, talvez o núcleo da equipe, contribuem para a coordenação, plano de ação, liderança, idéias estratégias e avaliação. 
    Existe também membros da equipe que fazem o trabalho necessário do dia-a-dia. Podemos chamá-los de implementadores, porque eles aceitam os papéis, restrições da organização e fazem a tarefa de produzir as coisas dentro do sistema.

    O ponto forte do implementador está na capacidade de traduzir conceitos gerais e planos em trabalhos práticos, e depois por em execução esses planos de trabalho de uma forma sistemática. Ele não está muito interessado em ideias entusiasmadas e inovações, possíveis opções ou estratégias. Procura introduzir, se necessário, objetivos claros, rotinas práticas de trabalho e resultados atingíveis. 
    Tendo alcançado isso, eles trabalham com profundidade, determinação e senso comum para alcançar os objetivos. Inversamente, ficam insatisfeitos e menos efetivos, onde procedimentos e objetivos não estão claramente explicados, onde flexibilidade, adaptabilidade e jogo de cintura são requeridos, junto com um rápido sistema de mudança e facilidade para cessar perdas. 
    Sua força de caráter pode significar que concorre para posições de liderança, e seu consistente desempenho anterior freqüentemente indica que são promovidos para níveis seniores. Sua relativa carência de visão e de capacidade de lidar com o instável e  ambientes de rápidas mudanças, aliado ao seu conservadorismo, pode significar que é menos efetivo como cabeça de um time do que como espinha dorsal.

    Frases e slogans que caracterizam um IMPLEMENTADOR:

    1. Se pode ser feito, nós faremos.
    2. Uma grama de ação vale mais do que um quilo de teoria.
    3. Trabalho duro nunca matou ninguém.
    4. Se for difícil, faremos já. Se for impossível, leva um pouco mais de tempo.
    5. Errar é humano, perdoar não é política da empresa.
    6. Vamos ao que interessa.
    7. A companhia tem todo o meu apoio.\n\n"""

msginvestigador = """
    INVESTIGADOR DE RECURSOS
    MANTER CONTATO COM OUTRAS EQUIPES

    PAPEL: Explorar e expor idéias, avanços e recursos fora do grupo; criar contatos externos que poderão ser úteis para a equipe e conduzir negociações subseqüentes.

    CARACTERÍSTICAS:
    • Pontos Fortes: uma personalidade descontraída, com forte senso inquisitivo e uma habilidade em enxergar possibilidades inerentes a qualquer coisa nova.
    • Pontos fracos toleráveis: entusiasmado demais e com carência de follow-up.
    • Habilidades comportamentais: produzir “centelhas” e networking.

    Descrição: Um notável membro de equipe, dominantemente extrovertido, de postura indagadora e inquieta com relação a vida. Ele é notável porque dificilmente é visto, por muito tempo, com os outros na sala, e quando se encontra, está sempre fazendo perguntas. 
    Sua força movedora de ação é explorar recursos e idéias fora da equipe e manter uma grande gama de contatos úteis. Quando está ‘em casa’, mantém o bom relacionamento dentro do grupo, como um formador de equipe, encorajando os colegas a usarem seus talentos e exporem suas idéias.

    Geralmente de caráter prestativo/alegre;  tem sorte por ter sua inclinação natural e forma de satisfação diretamente relacionada com sua maior contribuição para a equipe e é normalmente reconhecido como tal. 
    Existe alguns perigos: o investigador de recursos pode ser carente de auto-disciplina, ser impulsivo em seus interesses e de pronto abandonar um projeto. 
    Ele precisa de variedade, desafio, pessoas que constantemente os estimulem para mantê-lo satisfeito e produtivo. Diferente do inovador, não é um conjunto de idéias dentro de si mesmo, sua habilidade e contribuição consiste em estimular idéias dos outros, explorar novas possibilidades no mundo fora da equipe, e persuadir e motivar seus colegas. 
    Se perder o ímpeto, pode se tornar cansativo e ser desmoralizado. 
    Se ele perde o foco, pode perder tempo com coisas pouco relevantes e trivialidades. Contudo, sem ele, o time pode ficar olhando muito para dentro de si próprio, defensivo, cauteloso e sem contato com o resto do mundo. 

    Frases e slogans que caracterizam um INVESTIGADOR DE RECURSOS:

    1. Nós poderíamos fazer uma fortuna com isso.
    2. Ideias deveriam ser copiadas com orgulho.
    3. Nunca reinvente a roda.
    4. Oportunidades surgem com os erros dos outros.
    5. Certamente nós podemos tirar proveito disso.
    6. Você pode sempre telefonar para descobrir.
    7. Tempo gasto com reconhecimento raramente é desperdício.\n\n"""

msgformador = """
    FORMADOR DE EQUIPE
    BOM RELACIONAMENTO NA EQUIPE

    PAPEL: Apoiar os membros da equipe em seu pontos fortes (p.ex. fundamentando sugestões) e sustenta-los em suas deficiências; melhorar a comunicação, instigando o espírito de equipe em geral.

    CARACTERÍSTICAS:
    • Pontos fortes: humildade, flexibilidade, popularidade e boa habilidade de escuta.
    • Pontos Fracos Toleráveis: carência de firmeza e dureza, aversão a conflitos e competição, interrompe conflitos muito cedo.
    • Habilidades comportamentais: apoiar, extrair informações das pessoas, rastrear moral e sentimentos, lembrar aos outros sobre o lado humano.

    DESCRIÇÃO: Os cinco papéis - de coordenador, moldador, inovador, monitor/ avaliador e implementador - descrevem características  como poder, estrutura, status idéias e tarefas de alta prioridade para a execução dos mesmos. 
    Excetuando o coordenador, ‘pessoas’ como indivíduos não têm muita importância para esses  papéis, a não ser como complementos, ativos ou ‘ferramentas’ necessárias. Para o inovador, monitor / avaliador as pessoas podem ser uma interferência positiva para a beleza de simetria de seus pensamentos. Contudo, na maioria das equipes, existe pelo menos uma pessoa que não somente coloca as ‘pessoas’ praticamente no topo de seus interesses, como também possui a habilidade de usar esse seu interesse de forma positiva, mantendo as pessoas unidas. Com certa falta de imaginação, ele é chamado de formador de equipe.

    O formador de equipe é receptivo aos sentimentos, necessidades e interesses das pessoas do grupo. 
    Ele é um observador dos pontos fortes e fracos da equipe e é sua natureza ajudar a estimular os pontos fortes e sustentar os pontos fracos da equipe. 
    Os resultados são normalmente vistos, não somente em termos de espírito de equipe, mas também na comunicação, cooperação e resultados positivos em geral. 
    Talvez sua contribuição para a efetividade do grupo seja essencial e persuasiva, e emerge de forma mais nítida quando ele entra para minimizar os atritos que inovadores podem causar sem perceber e moldadores sem se importarem. 
    Ele poderá prevenir conflitos em potencial entre inovador e monitor ou entre moldador e implementador.

    Frases e slogans que caracterizam um FORMADOR DE EQUIPE:
    
    1. Cortesia nada custa.
    2. Eu estou bastante interessado no seu ponto de vista.
    3. Se está bom para vocês, está bom para mim.
    4. Todo mundo tem algo para contribuir.
    5. Se as pessoas se escutassem mais, falariam menos.
    6. Sempre se pode sentir uma atmosfera boa no trabalho.
    7. Eu tento ser versátil.\n\n"""  

msgfinalizador = """
    FINALIZADOR / ARREMATADOR
    MANTER A EQUIPE TRABALHANDO

    PAPEL: Assegurar que a equipe está protegida, tanto quanto possível, de erros de  omissão e incumbência; investigar ativamente aspectos do trabalho que necessitam de um maior grau de atenção e manter o senso de urgência dentro da equipe.

    CARACTERÍSTICAS:
    • Pontos fortes: uma habilidade em combinar um senso de interesse com um senso de ordem e propósito; auto controle e força de caráter. 
    • Pontos fracos toleráveis: impaciente e intolerante com relação àqueles que tem propensão e hábitos negligentes..
    • Habilidades comportamentais: rastrear detalhes, analisar e ordenar dados.

    DESCRIÇÃO: O papel do finalizador/arrematador tem características opostas a extroversão. Ele é altamente ansioso, freqüentemente compulsivo, tenso e introvertido. 
    Desenvolve uma energia consideravelmente tensa e usa esta energia de forma produtiva. Utiliza seus receios, medos e compulsões, e canaliza-os para o fechamento, em tempo, do trabalho. É chamado de finalizador.

    Sem o finalizador, um time comete erros ou  esquece coisas, atrasa na agenda de projetos e comete erros em pequenos detalhes que podem comprometer o trabalho final. 
    Pontos menos importantes podem ser deixados de fora ou esquecidos, mesmo que tenham se tornado urgente mais tarde, e a complacência pode se tornar parte do espírito da equipe . 
    O finalizador não deixa isto acontecer, ele se interessa, se preocupa, exige dos colegas, transmite um senso de urgência e é inimigo da preguiça.

    Um importante ingrediente da personalidade do finalizador é o auto controle, força de caráter e auto controle. Ele pode ficar nervoso, mas é um nervosismo bem canalizado, que vai direto ao ponto de uma forma energética e direta. 
    Obviamente, o finalizador não é a pessoa mais fácil de se conviver. 
    Corre-se o perigo dele baixar a moral da equipe, e se pegar envolvido demais em detalhes, levando outros companheiros da equipe a agirem da mesma forma. 
    Ele pode irritar e aborrecer as pessoas e faze-las ficarem tão nervosas quanto ele, mas não deixará que elas se tornem descuidadas ou confiantes em excesso. 
    Também não permitirá que percam tempo, quer seja com novos produtos quer seja com a solução de problema estratégicos. Quando aceito pela equipe, pode transformar o pensamento da equipe e contribuir imensamente para seu sucesso.     

    Frases e slogans que caracterizam um FINALIZADOR / ARREMATADOR:

    1. Isso é uma coisa que demanda total atenção.
    2. É preciso ler nas entrelinhas.
    3. Se alguma coisa tiver que dar errado, dará.
    4. Não existe desculpa para não ser perfeito.
    5. Somente a perfeição é suficiente.
    6. Antes que o mal cresça, corte-o pela raiz.
    7. Isso foi checado?\n\n""" 

msg = """
    Olá {},

        Segue o resultado do Teste de Belbin que você realizou. 
        O teste de Belbin mostra como uma pessoa trabalha em equipe.
        Assim, trata-se de uma ferramenta gerencial para autoconhecimento e conhecimento de outros. 

    Grande abraço,
    Alessandro Prudêncio Lukosevicius, 
    Professor, Consultor, Pesquisador e Escritor,
    Doutor, PMP, PRINCE2 Approved Trainer, MSP Practitioner\n\n"""

def ValidaRespostas (row):  
    return ((row['Id'] + row['If'] + row['Ic'] + row['Ia'] + row['Ib'] + row['Ig'] + row['Ih'] + row['Ie'] == 10) 
            and
            (row['IIb'] + row['IIe'] + row['IIg'] + row['IIc'] + row['IIf'] + row['IIa'] + row['IId'] + row['IIh'] == 10) 
            and 
            (row['IIIa'] + row['IIIc'] + row['IIId'] + row['IIIf'] + row['IIIe'] + row['IIIh'] + row['IIIg'] + row['IIIb'] == 10) 
            and
            (row['IVh'] + row['IVb'] + row['IVe'] + row['IVg'] + row['IVa'] + row['IVd'] + row['IVc'] + row['IVf'] == 10) 
            and
            (row['Vf'] + row['Vd'] + row['Vh'] + row['Ve'] + row['Vc'] + row['Vb'] + row['Va'] + row['Vg'] == 10) 
            and
            (row['VIc'] + row['VIg'] + row['VIa'] + row['VIh'] + row['VIb'] + row['VIf'] + row['VIe'] + row['VId'] == 10) 
            and 
            (row['VIIg'] + row['VIIa'] + row['VIIf'] + row['VIId'] + row['VIIh'] + row['VIIe'] + row['VIIb'] + row['VIIc'] == 10)) 
   
for i, row in df.iterrows(): #i: dataframe index; row: each row in series format
    if (row['data'].date().strftime('%d/%m/%Y') == dt.date.today().strftime('%d/%m/%Y')) and ValidaRespostas(row):
        coordenador   = row['Id'] + row['IIb'] + row['IIIa'] + row['IVh'] + row['Vf'] + row['VIc'] + row['VIIg'] 
        moldador      = row['If'] + row['IIe'] + row['IIIc'] + row['IVb'] + row['Vd'] + row['VIg'] + row['VIIa'] 
        inovador      = row['Ic'] + row['IIg'] + row['IIId'] + row['IVe'] + row['Vh'] + row['VIa'] + row['VIIf'] 
        investigador  = row['Ia'] + row['IIc'] + row['IIIf'] + row['IVg'] + row['Ve'] + row['VIh'] + row['VIId'] 
        formador      = row['Ib'] + row['IIf'] + row['IIIe'] + row['IVa'] + row['Vc'] + row['VIb'] + row['VIIh'] 
        implementador = row['Ig'] + row['IIa'] + row['IIIh'] + row['IVd'] + row['Vb'] + row['VIf'] + row['VIIe'] 
        monitor       = row['Ih'] + row['IId'] + row['IIIg'] + row['IVc'] + row['Va'] + row['VIe'] + row['VIIb'] 
        finalizador   = row['Ie'] + row['IIh'] + row['IIIb'] + row['IVf'] + row['Vg'] + row['VId'] + row['VIIc'] 
          
        list_perfis = [coordenador, moldador, inovador, monitor, implementador, investigador, formador, finalizador]
        list_nome_perfis = [msgcoordenador, msgmoldador, msginovador, msgmonitor, msgimplementador, msginvestigador, msgformador, msgfinalizador]
        list_pos = [i for i, item in enumerate(list_perfis) if item == max (list_perfis)]
        list_msg = [list_nome_perfis[pos] for pos in list_pos]

        mensagem = '==============================================================================================================\n\n'.join(list_msg)
        mensagem = msg.format(row['nome']) + mensagem
            
        EnviaEmailTodosAlunos("Resultado do Teste de Belbin", "alessanpl@gmail.com", "alessanpl@gmail.com", row['email'], mensagem)
        print('Mensagem enviada para {} com e-mail {}'.format(row['nome'].split()[0], row['email']))

Consmtp.quit()