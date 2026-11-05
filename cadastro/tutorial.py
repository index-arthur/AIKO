"""
Texto de ajuda do app, em um lugar so.

Ficava embutido no main.py e por isso envelheceu sem ninguem notar: nao
falava da aba Devolucao, que existe desde a v6.6. Separado aqui, as duas
cascas (Tk e WebView2) mostram o mesmo texto e ha um unico lugar para
atualizar quando algo muda.
"""

TEXTO = """INFORMAÇÕES NECESSÁRIAS:

• Usuário/Senha  → Seu login do TracKit (sem o @aiko.digital).

                   A senha é usada para autenticar e nunca vai para o
                   log. Se você marcar "Usar sempre neste PC", ela fica
                   guardada criptografada pelo Windows — só o seu usuário,
                   nessa máquina, consegue lê-la de volta. Nas próximas
                   vezes basta deixar o campo em branco. Para remover,
                   clique em "Esquecer senha salva"; para trocar, digite
                   outra por cima.

                   Clientes que entram por SSO (a tela com os botões
                   'Entrar com email da AIKO') não aceitam usuário e senha
                   aqui — o app detecta isso e abre o navegador para você
                   entrar por lá. A sessão capturada vale igual.

                   Depois do 1º login a sessão fica em cache e os campos
                   podem ficar vazios.

• Empresa        → Sigla da empresa (ex: BRC, RAI, QA.BRC)
• Equipamento    → Tipo (ex: COMODATO, SERVICO DE CAMPO)
• Ticket         → Escolha a origem no seletor: HWS (ClickUp) ou SP (SGI).
                   Cole o ticket inteiro (HWS-12312, SP-4471) ou só o
                   número. Maiúscula ou minúscula, com traço, espaço ou
                   sem nada — tudo funciona. Se você colar um ticket com
                   prefixo, o seletor acompanha sozinho.
• Zendesk        → Número do ticket Zendesk (deixe vazio se não tiver)
• Parou no bordo → Se parou em algum bordo, o número dele (0 se não parou)
• Qtd. Bordos    → Total de bordos a cadastrar


ABA CADASTRO DE BORDO:

1. Preencha Empresa (e usuário/senha, se a sessão expirou) e clique em
   Conectar.
2. Escolha Modelo, Grupo e Perfil nos seletores — eles listam os
   cadastros reais do cliente, com busca por nome ou ID.
3. Preencha os dados do lote. A prévia mostra como vai ficar o nome do
   primeiro equipamento e quantos serão criados.
4. Clique em Simular para conferir sem gravar nada.
5. Clique em Cadastrar para valer.

Cada equipamento criado é lido de volta e conferido (nome, modelo, perfil
e grupo). Divergência aparece no log em vermelho.

NOME LIVRE: no seletor "Nome", troque de Montado para Livre quando o lote
não seguir o padrão de ticket. Você escreve o nome inteiro e a automação
só acrescenta a numeração (| 01, | 02...). A numeração continua existindo
de propósito: é dela que a aba Vinculação depende para achar o lote.

COPIAR NOMES: o botão ao lado da prévia copia o lote inteiro, um nome por
linha, pronto para colar na planilha.


ABA VINCULAÇÃO:

Liga os computadores de bordo aos equipamentos. Informe o filtro
(ex: HWS-6848) e cole os seriais, um por linha.

O pareamento é POR POSIÇÃO: o 1º serial vai para o equipamento 01, o 2º
para o 02, e assim por diante.

Não precisa vincular o lote inteiro de uma vez. Cole 3 seriais de um
ticket de 10 e ele usa os equipamentos 01 a 03; depois cole os 7
restantes e ele continua do 04, porque pula quem já tem device. Só recusa
se sobrarem seriais sem equipamento livre.

Device que ainda não existe no TracKit é CADASTRADO e vinculado no mesmo
passo — a simulação marca esses com [NOVO].

Barra antes de gravar quando: o serial já está vinculado a outro
equipamento, está repetido na sua lista, está cadastrado em duplicidade
no TracKit, ou sobram seriais sem equipamento livre.

IMPORTANTE: Simular nunca grava. Confira o de-para no log e só então
clique em Vincular. Como device inexistente vira cadastro novo, um IMEI
digitado errado não dá erro — vira um device com o número torto. A
conferência na simulação é o que evita isso.


ABA STARLINK:

Cole os S/N dos kits, um por linha. Se quiser, o patrimônio vai na mesma
linha, ao lado (colado do Excel já vem assim).

Marque onde gravar: TracKit, ClickUp ou os dois.

No ClickUp (board Ativos TI) o casamento é pelo S/N:
  • S/N que ainda não está em nenhum card → cria o próximo SM/SV da
    sequência;
  • S/N que já está em um card → atualiza aquele card.
Só os campos marcados são gravados; os outros ficam como estão. Com o
TracKit marcado junto, o Cadastro no Trackit vira SIM sozinho.

Na primeira vez, Conectar ClickUp pede o seu token pessoal (ClickUp >
avatar > Configurações > Apps > API Token). Ele fica criptografado neste
computador, só para o seu usuário.

O ClickUp só roda depois de um TracKit limpo: gravar no board um kit que
não entrou no TracKit criaria um cadastro que ninguém consegue casar
depois.


ABA DEVOLUÇÃO:

Para quando o equipamento volta do campo.

PROCURAR: digite os últimos caracteres do deviceID (ex: f861), o ticket,
o número do Zendesk ou parte do nome do equipamento — tudo cai na mesma
busca. Serve para não precisar transcrever os 16 caracteres do device na
mão. Dos resultados você pode copiar o deviceID, copiar o nome
"Reserva - <deviceID>" já montado, ou jogar na lista de baixo.

Equipamento SEM nenhum bordo também aparece na busca, marcado como
"(sem bordo)" — é o que sobra depois de uma devolução, e é assim que você
o alcança para excluir.

DESVINCULAR: cole os deviceID, um por linha, e escolha a ação:
  • Só desvincular → o bordo fica "Não associado", como na tela do
    TracKit. O cadastro do bordo NÃO é apagado: ele continua na base para
    ser vinculado de novo.
  • Desvincular e EXCLUIR o equipamento → faz o desvínculo, confere por
    leitura de volta e só então exclui. Exclusão não tem desfazer, por
    isso pede que você digite EXCLUIR.

A TRAVA DOS 30 DIAS: antes de gravar qualquer coisa, o app pergunta ao
TracKit quando cada equipamento comunicou pela última vez.
  • Comunicou nos últimos 30 dias → o lote PARA. A máquina está em campo
    e desvincular faria ela parar de guardar dados sem ninguém perceber.
  • Mudo há mais de 30 dias, ou nunca comunicou → segue.
  • Não deu para verificar → o lote PARA também. Não saber se a máquina
    está em campo pesa igual a saber que está.

A caixa "Desvincular mesmo os que ainda comunicam" existe para a exceção
real, e volta desmarcada a cada lote.


ATUALIZAÇÃO:

Quando sai versão nova, aparece um aviso no topo ao abrir o app. Clicar em
Atualizar baixa e instala sozinho; o app fecha e reabre. A checagem
acontece só na abertura — se você deixa o app aberto o dia todo, feche e
abra para ver se saiu versão nova.
"""
