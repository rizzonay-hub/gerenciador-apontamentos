# -*- coding: utf-8 -*-
"""Monta os vídeos de treinamento a partir das capturas, dos cartões e da narração.

Uso:  python montar.py 1      (Trilha 1)
      python montar.py 2      (Trilha 2)

Cada cena vira um clipe (imagem + zoom lento + narração) e os clipes são emendados.
Saída: Trilha1-Apontamentos-de-horas.mp4 / Trilha2-Gerenciamento-do-Time.mp4
"""
import io, json, os, subprocess, sys, textwrap
from PIL import Image, ImageDraw, ImageFont

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.dirname(AQUI)
FF = os.path.join(AQUI, 'ferramentas', 'ffmpeg-9.0.2-essentials_build', 'bin')
FFMPEG, FFPROBE = os.path.join(FF, 'ffmpeg.exe'), os.path.join(FF, 'ffprobe.exe')
FONTE = os.path.join(RAIZ, 'Neulis_Neue_Regular.otf')
LOGO = os.path.join(RAIZ, 'LOGO_PRINCIPAL_2.png')
CAP_FUNC = os.path.join(AQUI, 'capturas-funcionario')
CAP_GESTOR = os.path.join(RAIZ, 'capturas')

W, H = 1920, 1080
ESC = 2                      # as imagens de entrada são renderizadas em 2x para o zoom ficar suave
FPS = 25
PRETO, BRANCO = (17, 17, 17), (255, 255, 255)
VERMELHO = (211, 29, 44)
CINZA_FUNDO = (243, 244, 246)
CINZA_TEXTO = (90, 94, 102)
CINZA_CLARO = (190, 194, 200)
ALT_TOPO, ALT_RODAPE = 110, 64

# ---------- áreas de destaque (em pixels da própria captura) ----------
DESTAQUES = {
    ('01-apontamentos', 'cartoes-hoje-mes'): (40, 245, 812, 446),
    ('01-apontamentos', 'cartao-menos-8h'): (1150, 245, 1600, 446),
    ('01-apontamentos', 'botao-novo'): (1600, 155, 1905, 222),
    ('01-apontamentos', 'botao-excluir'): (1282, 155, 1446, 222),
    ('01-apontamentos', 'botao-exportar'): (1440, 155, 1607, 222),
    ('01-apontamentos', 'botoes-gestor'): (860, 155, 1718, 286),
    ('01-apontamentos', 'botao-avisos'): (1535, 222, 1718, 286),
    ('02-novo-apontamento', 'bloco-o-que'): (34, 194, 1142, 452),
    ('02-novo-apontamento', 'bloco-quando'): (34, 462, 1142, 752),
    ('02-novo-apontamento', 'atalhos'): (56, 610, 692, 702),
    ('02-novo-apontamento', 'botoes-salvar'): (708, 788, 1145, 858),
    ('02-novo-apontamento', 'campo-colaborador'): (662, 532, 1112, 606),
    ('02b-busca-cliente', 'busca'): (54, 296, 972, 568),
    ('02c-acima-de-8h', 'subtitulo'): (34, 108, 620, 152),
    ('03-varias-empresas', 'caixa-varias'): (54, 388, 345, 430),
    ('03-varias-empresas', 'listas'): (54, 492, 1118, 908),
    ('03-varias-empresas', 'campo-grupo'): (54, 432, 878, 490),
    ('03b-dividir-tempo', 'como-gravar'): (54, 914, 730, 994),
    ('08-gerenciamento-analitica', 'abas'): (30, 142, 830, 210),
    ('12-gerenciamento-avisos', 'grade'): (6, 380, 1876, 545),
    ('12-gerenciamento-avisos', 'marcar-vistos'): (346, 314, 666, 384),
}

# ---------- fontes ----------
_fontes = {}
def fonte(tam):
    if tam not in _fontes:
        _fontes[tam] = ImageFont.truetype(FONTE, tam)
    return _fontes[tam]

def duracao(mp3):
    out = subprocess.check_output([FFPROBE, '-v', 'error', '-show_entries', 'format=duration', '-of', 'csv=p=0', mp3])
    return float(out.decode().strip())

def quebrar(draw, texto, f, largura):
    palavras, linhas, atual = texto.split(' '), [], ''
    for p in palavras:
        teste = (atual + ' ' + p).strip()
        if draw.textlength(teste, font=f) <= largura:
            atual = teste
        else:
            if atual: linhas.append(atual)
            atual = p
    if atual: linhas.append(atual)
    return linhas

# ---------- elementos comuns ----------
def logo_redimensionado(largura):
    im = Image.open(LOGO).convert('RGBA')
    esc = largura / im.width
    return im.resize((int(im.width * esc), int(im.height * esc)), Image.LANCZOS)

def cabecalho(img, draw, curso, direita=None):
    e = ESC
    draw.rectangle([0, 0, W * e, ALT_TOPO * e], fill=PRETO)
    lg = logo_redimensionado(96 * e)
    img.paste(lg, (36 * e, (ALT_TOPO * e - lg.height) // 2 + 4 * e), lg)
    draw.text((160 * e, 32 * e), curso, font=fonte(34 * e), fill=BRANCO)
    if direita:
        f = fonte(24 * e)
        draw.text((W * e - 48 * e - draw.textlength(direita, font=f), 42 * e), direita, font=f, fill=CINZA_CLARO)

def rodape(draw, esquerda, direita):
    e = ESC
    y = H * e - ALT_RODAPE * e
    draw.rectangle([0, y, W * e, H * e], fill=BRANCO)
    draw.line([0, y, W * e, y], fill=(225, 227, 231), width=2 * e)
    f = fonte(20 * e)
    draw.text((48 * e, y + 20 * e), esquerda, font=f, fill=CINZA_TEXTO)
    draw.text((W * e - 48 * e - draw.textlength(direita, font=f), y + 20 * e), direita, font=f, fill=CINZA_TEXTO)

def novo_canvas(cor):
    img = Image.new('RGB', (W * ESC, H * ESC), cor)
    return img, ImageDraw.Draw(img)

# ---------- cartões ----------
def cartao_escuro(titulo, subtitulo=None, etiqueta=None, com_logo=True):
    img, d = novo_canvas(PRETO)
    e = ESC
    y = 300 * e
    if com_logo:
        lg = logo_redimensionado(420 * e)
        img.paste(lg, ((W * e - lg.width) // 2, 230 * e), lg)
        y = 230 * e + lg.height + 40 * e
    if etiqueta:
        f = fonte(30 * e)
        d.text(((W * e - d.textlength(etiqueta, font=f)) // 2, y), etiqueta, font=f, fill=VERMELHO)
        y += 60 * e
    f = fonte(72 * e)
    d.text(((W * e - d.textlength(titulo, font=f)) // 2, y), titulo, font=f, fill=BRANCO)
    y += 100 * e
    d.rectangle([(W * e - 120 * e) // 2, y, (W * e + 120 * e) // 2, y + 6 * e], fill=VERMELHO)
    y += 40 * e
    if subtitulo:
        f = fonte(36 * e)
        d.text(((W * e - d.textlength(subtitulo, font=f)) // 2, y), subtitulo, font=f, fill=CINZA_CLARO)
    return img

def cartao_texto(curso, etiqueta, linhas, mostrar=None, rodape_esq='', rodape_dir='', grande=None):
    """Cartão branco com etiqueta vermelha e linhas com marcador. `mostrar` limita quantas linhas aparecem."""
    img, d = novo_canvas(BRANCO)
    e = ESC
    cabecalho(img, d, curso, rodape_dir)
    y = 200 * e
    f = fonte(30 * e)
    d.text((120 * e, y), etiqueta.upper(), font=f, fill=VERMELHO)
    y += 70 * e
    if grande:
        f = fonte(66 * e)
        for l in quebrar(d, grande, f, (W - 240) * e):
            d.text((120 * e, y), l, font=f, fill=PRETO)
            y += 90 * e
    n = len(linhas) if mostrar is None else mostrar
    fl = fonte(46 * e)
    for i, linha in enumerate(linhas[:n]):
        d.rectangle([120 * e, y + 18 * e, 120 * e + 18 * e, y + 36 * e], fill=VERMELHO)
        for j, l in enumerate(quebrar(d, linha, fl, (W - 320) * e)):
            d.text((168 * e, y), l, font=fl, fill=PRETO)
            y += 62 * e
        y += 30 * e
    rodape(d, rodape_esq, rodape_dir)
    return img

# ---------- tela com captura ----------
def tela(curso, captura, destaque, rodape_esq, rodape_dir):
    """Devolve (imagem, caixa do destaque em pixels do canvas 1x ou None)."""
    img, d = novo_canvas(CINZA_FUNDO)
    e = ESC
    cabecalho(img, d, curso, rodape_dir)
    cap = Image.open(captura).convert('RGB')
    area_x0, area_y0 = 40, ALT_TOPO + 22
    area_w, area_h = W - 80, H - ALT_TOPO - ALT_RODAPE - 44
    esc = min(area_w / cap.width, area_h / cap.height)
    cw, ch = int(cap.width * esc), int(cap.height * esc)
    cx0 = area_x0 + (area_w - cw) // 2
    cy0 = area_y0 + (area_h - ch) // 2
    # sombra leve
    d.rectangle([(cx0 + 6) * e, (cy0 + 8) * e, (cx0 + cw + 6) * e, (cy0 + ch + 8) * e], fill=(214, 216, 220))
    img.paste(cap.resize((cw * e, ch * e), Image.LANCZOS), (cx0 * e, cy0 * e))
    caixa = None
    nome = os.path.splitext(os.path.basename(captura))[0]
    box = DESTAQUES.get((nome, destaque))
    if box:
        x0, y0, x1, y1 = box
        bx0, by0 = cx0 + x0 * esc, cy0 + y0 * esc
        bx1, by1 = cx0 + x1 * esc, cy0 + y1 * esc
        d.rectangle([bx0 * e, by0 * e, bx1 * e, by1 * e], outline=VERMELHO, width=5 * e)
        d.rectangle([(bx0 - 3) * e, (by0 - 3) * e, (bx1 + 3) * e, (by1 + 3) * e], outline=(255, 235, 237), width=1 * e)
        caixa = (bx0, by0, bx1, by1)
    rodape(d, rodape_esq, rodape_dir)
    return img, caixa

# ---------- clipe ----------
def clipe(frame_png, dur, saida, caixa=None, zoom_suave=True):
    """Imagem -> vídeo com zoom lento. `caixa` (x0,y0,x1,y1 no canvas 1x) é o centro do zoom."""
    frames = max(2, int(round(dur * FPS)))
    iw, ih = W * ESC, H * ESC
    if caixa:
        x0, y0, x1, y1 = caixa
        bw, bh = max(x1 - x0, 1), max(y1 - y0, 1)
        z = min(2.0, max(1.15, 0.62 * min(W / bw, H / bh)))
        cx, cy = ((x0 + x1) / 2) * ESC, ((y0 + y1) / 2) * ESC
    elif zoom_suave:
        z, cx, cy = 1.06, iw / 2, ih / 2
    else:
        z, cx, cy = 1.0, iw / 2, ih / 2
    # zoom entra nos primeiros 70% da cena, com curva suave
    zexp = f"1+({z:.4f}-1)*(1-cos(min(on/({frames}*0.7),1)*PI))/2"
    xexp = f"{cx:.1f}-(iw/zoom)/2"
    yexp = f"{cy:.1f}-(ih/zoom)/2"
    vf = (f"zoompan=z='{zexp}':x='{xexp}':y='{yexp}':d={frames}:s={W}x{H}:fps={FPS},"
          f"fade=t=in:st=0:d=0.25,fade=t=out:st={max(dur - 0.25, 0):.2f}:d=0.25,format=yuv420p")
    subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-i', frame_png, '-vf', vf, '-t', f'{dur:.3f}',
                           '-r', str(FPS), '-c:v', 'libx264', '-preset', 'medium', '-crf', '19', saida])

def juntar_audio(video, mp3, saida, atraso_ms=500):
    if mp3:
        subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-i', video, '-i', mp3,
                               '-filter_complex', f'[1:a]adelay={atraso_ms}:all=1,apad[a]',
                               '-map', '0:v', '-map', '[a]', '-c:v', 'copy', '-c:a', 'aac', '-b:a', '160k', '-ar', '44100', '-ac', '2',
                               '-shortest', saida])
    else:
        subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-i', video, '-f', 'lavfi', '-i', 'anullsrc=r=44100:cl=stereo',
                               '-map', '0:v', '-map', '1:a', '-c:v', 'copy', '-c:a', 'aac', '-b:a', '160k', '-shortest', saida])

def concatenar(clipes, saida):
    lista = os.path.join(os.path.dirname(clipes[0]), 'lista.txt')
    with io.open(lista, 'w', encoding='utf-8') as f:
        for c in clipes:
            f.write("file '" + c.replace("\\", "/").replace("'", "'\\''") + "'\n")
    subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-f', 'concat', '-safe', '0', '-i', lista, '-c', 'copy', '-movflags', '+faststart', saida])

# ---------- montagem de uma trilha ----------
def montar(trilha, cenas, curso, narracao, saida, total_partes):
    trabalho = os.path.join(AQUI, 'trabalho', f'trilha{trilha}')
    os.makedirs(trabalho, exist_ok=True)
    clipes = []
    for i, c in enumerate(cenas):
        cid = c['id']
        mp3 = os.path.join(narracao, f'{cid}.mp3')
        if not os.path.exists(mp3): mp3 = None
        dur_audio = duracao(mp3) if mp3 else 3.0
        parte = c.get('parte')
        rod_dir = f'Parte {parte} de {total_partes}' if parte else ''
        rod_esq = curso + '  ·  Treinamento'
        tipo = c['tipo']
        base = os.path.join(trabalho, f'{cid}')
        print(f'[{i+1}/{len(cenas)}] {cid} {tipo} {dur_audio:.1f}s', flush=True)

        if tipo == 'tela':
            pasta = CAP_GESTOR if c.get('gestor') else CAP_FUNC
            img, caixa = tela(curso, os.path.join(pasta, c['captura']), c.get('destaque', 'tela'), rod_esq, rod_dir)
            img.save(base + '.png')
            dur = dur_audio + 1.3
            clipe(base + '.png', dur, base + '_v.mp4', caixa)
            juntar_audio(base + '_v.mp4', mp3, base + '.mp4')
            clipes.append(base + '.mp4')

        elif tipo in ('abertura', 'encerramento', 'parte'):
            if tipo == 'parte':
                img = cartao_escuro(c['titulo'], etiqueta=f'PARTE {parte} DE {total_partes}', com_logo=False)
            else:
                img = cartao_escuro(c['titulo'], c.get('subtitulo'))
            img.save(base + '.png')
            dur = dur_audio + (1.6 if tipo != 'parte' else 1.0)
            clipe(base + '.png', dur, base + '_v.mp4', None, zoom_suave=False)
            juntar_audio(base + '_v.mp4', mp3, base + '.mp4')
            clipes.append(base + '.mp4')

        elif tipo == 'cartao':
            linhas = c.get('linhas', [])
            grande = c.get('grande')
            if linhas and c.get('uma_a_uma', True) and len(linhas) > 1:
                # as linhas aparecem uma a uma, no ritmo da narração
                n = len(linhas)
                total = dur_audio + 1.5
                partes = []
                for k in range(1, n + 1):
                    img = cartao_texto(curso, c['etiqueta'], linhas, k, rod_esq, rod_dir, grande)
                    png = f'{base}_{k}.png'; img.save(png)
                    d = total / n
                    sub = f'{base}_{k}_v.mp4'
                    frames = max(2, int(round(d * FPS)))
                    vf = f"zoompan=z='1':d={frames}:s={W}x{H}:fps={FPS},format=yuv420p"
                    subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-i', png, '-vf', vf, '-t', f'{d:.3f}', '-r', str(FPS),
                                           '-c:v', 'libx264', '-preset', 'medium', '-crf', '19', sub])
                    partes.append(sub)
                lista = f'{base}_lista.txt'
                with io.open(lista, 'w', encoding='utf-8') as f:
                    for p in partes: f.write("file '" + p.replace("\\", "/") + "'\n")
                subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-f', 'concat', '-safe', '0', '-i', lista, '-c', 'copy', base + '_v0.mp4'])
                # fade de entrada e saída no conjunto
                subprocess.check_call([FFMPEG, '-y', '-v', 'error', '-i', base + '_v0.mp4', '-vf',
                                       f'fade=t=in:st=0:d=0.25,fade=t=out:st={total-0.25:.2f}:d=0.25',
                                       '-c:v', 'libx264', '-preset', 'medium', '-crf', '19', base + '_v.mp4'])
            else:
                img = cartao_texto(curso, c['etiqueta'], linhas, None, rod_esq, rod_dir, grande)
                img.save(base + '.png')
                clipe(base + '.png', dur_audio + 1.5, base + '_v.mp4', None, zoom_suave=False)
            juntar_audio(base + '_v.mp4', mp3, base + '.mp4')
            clipes.append(base + '.mp4')
        else:
            raise ValueError(tipo)
    concatenar(clipes, saida)
    print('pronto:', saida, flush=True)

# ---------- roteiros ----------
T1_CURSO = 'Apontamentos de horas'
T1 = [
    {'id': '01', 'tipo': 'abertura', 'titulo': 'Apontamentos de horas', 'subtitulo': 'Treinamento para colaboradores  ·  6 minutos'},
    {'id': '02', 'tipo': 'cartao', 'etiqueta': 'O que você vai aprender', 'linhas': ['Marcar as horas do dia', 'Marcar para várias empresas de uma vez', 'Corrigir e completar um dia', 'O que acontece acima de 8 horas']},
    {'id': '03', 'tipo': 'cartao', 'etiqueta': 'Por que isso importa', 'linhas': ['Mostra o trabalho de cada pessoa', 'Mostra quanto cada cliente custa', 'Marque todos os dias, antes de ir embora']},
    {'id': '04', 'tipo': 'parte', 'parte': 1, 'titulo': 'A tela'},
    {'id': '05', 'tipo': 'cartao', 'parte': 1, 'etiqueta': 'Onde abrir', 'grande': 'Menu Time  →  Apontamentos', 'linhas': []},
    {'id': '06', 'tipo': 'tela', 'parte': 1, 'captura': '01-apontamentos.png', 'destaque': 'tela'},
    {'id': '07', 'tipo': 'tela', 'parte': 1, 'captura': '01-apontamentos.png', 'destaque': 'cartoes-hoje-mes'},
    {'id': '08', 'tipo': 'tela', 'parte': 1, 'captura': '01-apontamentos.png', 'destaque': 'cartao-menos-8h'},
    {'id': '09', 'tipo': 'cartao', 'parte': 1, 'etiqueta': 'O que lembrar', 'linhas': ['A tela tem três partes', 'O cartão Hoje mostra o seu dia', 'O cartão Menos de 8 horas mostra o que faltou']},
    {'id': '10', 'tipo': 'parte', 'parte': 2, 'titulo': 'Marcar as horas'},
    {'id': '11', 'tipo': 'tela', 'parte': 2, 'captura': '01-apontamentos.png', 'destaque': 'botao-novo'},
    {'id': '12', 'tipo': 'tela', 'parte': 2, 'captura': '02-novo-apontamento.png', 'destaque': 'bloco-o-que'},
    {'id': '13', 'tipo': 'tela', 'parte': 2, 'captura': '02b-busca-cliente.png', 'destaque': 'busca'},
    {'id': '14', 'tipo': 'tela', 'parte': 2, 'captura': '02-novo-apontamento.png', 'destaque': 'bloco-quando'},
    {'id': '15', 'tipo': 'tela', 'parte': 2, 'captura': '02-novo-apontamento.png', 'destaque': 'atalhos'},
    {'id': '16', 'tipo': 'tela', 'parte': 2, 'captura': '02-novo-apontamento.png', 'destaque': 'botoes-salvar'},
    {'id': '17', 'tipo': 'cartao', 'parte': 2, 'etiqueta': 'O que lembrar', 'linhas': ['Novo apontamento abre a janela', 'Em cima: o que você fez', 'Embaixo: quando e quanto tempo', 'Salvar guarda']},
    {'id': '18', 'tipo': 'parte', 'parte': 3, 'titulo': 'Várias empresas e grupos'},
    {'id': '19', 'tipo': 'tela', 'parte': 3, 'captura': '03-varias-empresas.png', 'destaque': 'caixa-varias'},
    {'id': '20', 'tipo': 'tela', 'parte': 3, 'captura': '03-varias-empresas.png', 'destaque': 'listas'},
    {'id': '21', 'tipo': 'tela', 'parte': 3, 'captura': '03b-dividir-tempo.png', 'destaque': 'como-gravar'},
    {'id': '22', 'tipo': 'tela', 'parte': 3, 'captura': '03c-meus-grupos.png', 'destaque': 'tela'},
    {'id': '23', 'tipo': 'tela', 'parte': 3, 'captura': '03-varias-empresas.png', 'destaque': 'campo-grupo'},
    {'id': '24', 'tipo': 'cartao', 'parte': 3, 'etiqueta': 'O que lembrar', 'linhas': ['Várias empresas: o mesmo trabalho, de uma vez', 'Tempo inteiro ou dividido: o que você gastou de verdade', 'Grupos: as empresas de sempre']},
    {'id': '25', 'tipo': 'parte', 'parte': 4, 'titulo': 'Corrigir e completar'},
    {'id': '26', 'tipo': 'tela', 'parte': 4, 'captura': '04-editar-apontamento.png', 'destaque': 'tela'},
    {'id': '27', 'tipo': 'tela', 'parte': 4, 'captura': '01-apontamentos.png', 'destaque': 'botao-excluir'},
    {'id': '28', 'tipo': 'tela', 'parte': 4, 'captura': '05-lancar-dia-pendente.png', 'destaque': 'tela'},
    {'id': '29', 'tipo': 'tela', 'parte': 4, 'captura': '01-apontamentos.png', 'destaque': 'botao-exportar'},
    {'id': '30', 'tipo': 'cartao', 'parte': 4, 'etiqueta': 'O que lembrar', 'linhas': ['Dois cliques na linha: corrige', 'Excluir: apaga', 'Clique no dia do cartão: completa']},
    {'id': '31', 'tipo': 'parte', 'parte': 5, 'titulo': 'Mais de 8 horas no dia'},
    {'id': '32', 'tipo': 'cartao', 'parte': 5, 'etiqueta': 'Mais de 8 horas no dia', 'linhas': ['O dia normal tem 8 horas', 'Passou de 8? O sistema guarda e avisa', 'O aviso vai para você e para o seu gestor']},
    {'id': '33', 'tipo': 'tela', 'parte': 5, 'captura': '02c-acima-de-8h.png', 'destaque': 'subtitulo'},
    {'id': '34', 'tipo': 'cartao', 'parte': 5, 'etiqueta': 'O que lembrar', 'linhas': ['Nada é bloqueado', 'Acima de 8 horas, você e o seu gestor são avisados', 'Marque o tempo de verdade']},
    {'id': '35', 'tipo': 'cartao', 'etiqueta': 'Resumo', 'linhas': ['Marque as horas todos os dias', 'Novo apontamento: cliente ou atividade interna', 'Várias empresas e Grupos: o mesmo trabalho de uma vez', 'Dois cliques corrige; o cartão completa', 'Mais de 8 horas: guarda e avisa']},
    {'id': '36', 'tipo': 'cartao', 'etiqueta': 'Dúvidas', 'linhas': ['Como usar: pergunte ao seu gestor', 'O sistema deu erro: fale com o Emerson']},
    {'id': '37', 'tipo': 'cartao', 'etiqueta': 'Agora é com você', 'linhas': ['Responda o formulário com 5 perguntas', 'Leva 2 minutos', 'O link está junto com este vídeo']},
    {'id': '38', 'tipo': 'encerramento', 'titulo': 'Treinamento concluído', 'subtitulo': 'Apontamentos de horas'},
]

T2_CURSO = 'Gerenciamento do Time'
T2 = [
    {'id': '01', 'tipo': 'abertura', 'titulo': 'Gerenciamento do Time', 'subtitulo': 'Treinamento para gestores  ·  4 minutos'},
    {'id': '02', 'tipo': 'cartao', 'etiqueta': 'O que você vai aprender', 'linhas': ['Marcar horas para alguém da equipe', 'Marcar vários dias de uma vez', 'Ver as horas da equipe', 'Cuidar dos avisos']},
    {'id': '03', 'tipo': 'cartao', 'etiqueta': 'Por que isso importa', 'linhas': ['Mostra quem está com trabalho demais', 'Mostra quanto cada cliente custa', 'O aviso é para conversar, não para punir']},
    {'id': '04', 'tipo': 'parte', 'parte': 1, 'titulo': 'Marcar pela equipe'},
    {'id': '05', 'tipo': 'tela', 'parte': 1, 'gestor': True, 'captura': '01-apontamentos.png', 'destaque': 'botoes-gestor'},
    {'id': '06', 'tipo': 'tela', 'parte': 1, 'gestor': True, 'captura': '02-novo-apontamento.png', 'destaque': 'campo-colaborador'},
    {'id': '07', 'tipo': 'tela', 'parte': 1, 'gestor': True, 'captura': '07-anotacoes-periodo.png', 'destaque': 'tela'},
    {'id': '08', 'tipo': 'cartao', 'parte': 1, 'etiqueta': 'O que lembrar', 'linhas': ['Campo Colaborador: marca para outra pessoa', 'Lançar período: vários dias de uma vez']},
    {'id': '09', 'tipo': 'parte', 'parte': 2, 'titulo': 'Ver as horas da equipe'},
    {'id': '10', 'tipo': 'tela', 'parte': 2, 'gestor': True, 'captura': '08-gerenciamento-analitica.png', 'destaque': 'abas'},
    {'id': '11', 'tipo': 'tela', 'parte': 2, 'gestor': True, 'captura': '09-gerenciamento-sintetica.png', 'destaque': 'tela'},
    {'id': '12', 'tipo': 'tela', 'parte': 2, 'gestor': True, 'captura': '10-gerenciamento-por-funcionario.png', 'destaque': 'tela'},
    {'id': '13', 'tipo': 'tela', 'parte': 2, 'gestor': True, 'captura': '11-gerenciamento-por-usuario-data.png', 'destaque': 'tela'},
    {'id': '14', 'tipo': 'tela', 'parte': 2, 'gestor': True, 'captura': '01-apontamentos.png', 'destaque': 'botao-exportar'},
    {'id': '15', 'tipo': 'cartao', 'parte': 2, 'etiqueta': 'O que lembrar', 'linhas': ['Analítica: cada marcação', 'Sintética: as somas', 'Por funcionário: cliente ou interno', 'Por usuário e data: quem está fora do esperado']},
    {'id': '16', 'tipo': 'parte', 'parte': 3, 'titulo': 'Avisos de horas a mais'},
    {'id': '17', 'tipo': 'cartao', 'parte': 3, 'etiqueta': 'Avisos de horas a mais', 'linhas': ['Passou de 8 horas: o sistema guarda e avisa', 'O aviso vai para o líder da equipe', 'Sem líder: vai para o gestor do departamento']},
    {'id': '18', 'tipo': 'tela', 'parte': 3, 'gestor': True, 'captura': '01-apontamentos.png', 'destaque': 'botao-avisos'},
    {'id': '19', 'tipo': 'tela', 'parte': 3, 'gestor': True, 'captura': '12-gerenciamento-avisos.png', 'destaque': 'grade'},
    {'id': '20', 'tipo': 'tela', 'parte': 3, 'gestor': True, 'captura': '12-gerenciamento-avisos.png', 'destaque': 'marcar-vistos'},
    {'id': '21', 'tipo': 'cartao', 'parte': 3, 'etiqueta': 'O que lembrar', 'linhas': ['O aviso é para conversar com a pessoa', 'Olhe os avisos e marque como vistos']},
    {'id': '22', 'tipo': 'cartao', 'etiqueta': 'Resumo', 'linhas': ['Campo Colaborador e Lançar período: marcar pela equipe', 'Gerenciamento: ver as horas', 'Avisos: olhar e marcar como vistos']},
    {'id': '23', 'tipo': 'cartao', 'etiqueta': 'Dúvidas', 'linhas': ['O sistema deu erro: fale com o Emerson']},
    {'id': '24', 'tipo': 'cartao', 'etiqueta': 'Agora é com você', 'linhas': ['Responda o formulário com 3 perguntas', 'O link está junto com este vídeo']},
    {'id': '25', 'tipo': 'encerramento', 'titulo': 'Treinamento concluído', 'subtitulo': 'Gerenciamento do Time'},
]

if __name__ == '__main__':
    qual = sys.argv[1] if len(sys.argv) > 1 else '1'
    if qual == '1':
        montar(1, T1, T1_CURSO, os.path.join(AQUI, 'narracao-trilha1'), os.path.join(AQUI, 'Trilha1-Apontamentos-de-horas.mp4'), 5)
    elif qual == '2':
        montar(2, T2, T2_CURSO, os.path.join(AQUI, 'narracao-trilha2'), os.path.join(AQUI, 'Trilha2-Gerenciamento-do-Time.mp4'), 3)
    elif qual == 'teste':
        # só os cartões e uma tela, para conferir o visual antes de montar tudo
        saida = os.path.join(AQUI, 'trabalho', 'teste')
        os.makedirs(saida, exist_ok=True)
        cartao_escuro('Apontamentos de horas', 'Treinamento para colaboradores  ·  6 minutos').resize((W, H), Image.LANCZOS).save(os.path.join(saida, 'abertura.png'))
        cartao_escuro('Marcar as horas', etiqueta='PARTE 2 DE 5', com_logo=False).resize((W, H), Image.LANCZOS).save(os.path.join(saida, 'parte.png'))
        cartao_texto(T1_CURSO, 'O que você vai aprender', T1[1]['linhas'], 3, T1_CURSO + '  ·  Treinamento', '').resize((W, H), Image.LANCZOS).save(os.path.join(saida, 'objetivos.png'))
        cartao_texto(T1_CURSO, 'Onde abrir', [], None, T1_CURSO + '  ·  Treinamento', 'Parte 1 de 5', 'Menu Time  →  Apontamentos').resize((W, H), Image.LANCZOS).save(os.path.join(saida, 'menu.png'))
        img, caixa = tela(T1_CURSO, os.path.join(CAP_FUNC, '01-apontamentos.png'), 'cartao-menos-8h', T1_CURSO + '  ·  Treinamento', 'Parte 1 de 5')
        img.resize((W, H), Image.LANCZOS).save(os.path.join(saida, 'tela.png')); print('caixa', caixa)
        img, caixa = tela(T1_CURSO, os.path.join(CAP_FUNC, '02-novo-apontamento.png'), 'atalhos', T1_CURSO + '  ·  Treinamento', 'Parte 2 de 5')
        img.resize((W, H), Image.LANCZOS).save(os.path.join(saida, 'tela2.png')); print('caixa', caixa)
        print('teste ok')
