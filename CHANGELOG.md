# Histórico — Apontamentos de horas

## 16/10/2026 (tarde) — hierarquia do menu Time dentro da tela

**Novo**
- **Lançar período** (`frmApontamentoPeriodo`): o "Anotações Período" de hoje no visual novo. Gestor ou diretoria escolhe a pessoa, o intervalo e o motivo; a janela mostra antes quantos dias e horas vai gravar (8h00 por dia útil, "Sabado" 4h00, domingo nada, como o `frmTimePeriodo`) e pula os dias em que a pessoa já tem apontamento.
- **Gerenciamento do Time** (`frmGerenciamentoTime`): uma tela com uma aba por consulta de gestão. A **Analítica** está pronta, com as mesmas procedures `Time_GetByEmpresa*` e a mesma escolha entre elas que o `frmTimeConsulta`; Sintética, Por funcionário e Por usuário e data aparecem como "em breve".
- Dois botões na barra da lista, só para gestor e diretoria: "Lançar período" e "Gerenciamento".
- `ModApontamento`: `ListarDe`, `Colaboradores` (TbSenha), `ConsultaAnalitica`, `PodeGerenciar`. Demo: `colaboradores.txt` fictício e lançamentos de mais duas pessoas; a pessoa logada na demo passou a ser gestor.

**Mudou**
- Banco de horas: sábado conta 4h00 de jornada (antes 0), igual à meta mensal e ao "Sabado" do período; só o que passar disso é crédito.

## 16/10/2026 — pronta para pré-produção

**Novo**
- **Rateio entre empresas**: em "Lançar para várias empresas", a opção "Como gravar o tempo?" com *cada empresa fica com o tempo inteiro* (padrão, como o Time de hoje) ou *dividir o tempo entre as empresas*. Divisão em centésimos de hora para a soma fechar exata; recusa quando não dá 0,01 h por empresa. As frases mostram os números reais da pessoa.
- **Trava de 10 h por dia** (`BancoHoras.LIMITE_DIA`): o que já está no dia mais o que vai ser gravado não passa de 10h00. Vale para novo, alteração e várias empresas. O subtítulo da janela mostra o efeito antes de gravar.
- **Banco de horas** (`BancoHoras`, em `Apontamento.vb`): o que passa de 8h00 num dia útil (e todo o tempo de sábado e domingo) é crédito, desde 01/10/2026. Calculado a partir da `TbTime`, sem tabela nova. Quinto cartão BANCO DE HORAS na tela e extrato `frmBancoHoras` com exportação CSV.
- Demo: dados gerados com dois dias acima de 8 h; captura `06-banco-de-horas.png`.

**Corrigido**
- Cartão MENOS DE 8 H: o "+N" dos dias que não cabiam era desenhado em cima do aviso do rodapé.
- Rodapé da janela de lançamento: o texto de atalhos passava por baixo do botão Cancelar.
- "1 apontamentos gravados" quando só uma empresa estava na lista.

**Visual e textos**
- Números grandes dos cartões em Segoe UI negrito (a fonte da marca, Neulis Neue, fica só nos títulos). O "h" segue minúsculo.
- Textos das opções e avisos reescritos em frases curtas, com os números da pessoa, sem decimais na tela.

**Documentação**
- `README.md`, `IMPLANTACAO.md` (passo a passo e roteiro de teste de pré-produção), `CHANGELOG.md` e `.gitignore`.

## 25/09/2026 — primeira entrega (demonstração)

- Tela `frmApontamentos` no visual novo: período, busca, quatro cartões (HOJE, ESTE MÊS com gráfico, MÊS ANTERIOR, MENOS DE 8 H), lista, exportar CSV.
- Janela `frmApontamentoDet`: cliente com filtro conforme digita, atividade interna, tempo em horas e minutos com atalhos, "Salvar e novo", várias empresas (mesmo tempo para cada).
- `ResumoHoras` reúne a regra da trava de horas do `frmMain` num lugar só.
- `ModApontamento` (SQL, mesmas procedures do Time) e `ModApontamentoDemo` (txt, sem banco).
- Saíram: campo Retrabalho (oculto desde 2022), CNPJ digitável (era só exibição), o "Inserir antes de preencher".
- Demo `GM-Apontamentos-Demo.exe` e a mesma tela embutida na demo da tela principal (v12).
