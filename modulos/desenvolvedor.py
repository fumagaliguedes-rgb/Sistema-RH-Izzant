import os
import sys
import sqlite3
import traceback
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser


def _status_text(ok):
    return 'OK' if ok else 'ERRO'


def _write_log(log_dir, texto):
    os.makedirs(log_dir, exist_ok=True)
    path = os.path.join(log_dir, 'diagnostico.log')
    with open(path, 'a', encoding='utf-8') as f:
        f.write('\n' + '='*80 + '\n')
        f.write(datetime.now().strftime('%d/%m/%Y %H:%M:%S') + '\n')
        f.write(texto + '\n')
    return path


def build_modo_desenvolvedor(app, parent, ctx):
    """Tela de diagnóstico e homologação do Sistema RH Izzant.

    Este módulo não altera a folha de ponto aprovada. Ele apenas valida o ambiente,
    pastas, banco, dependências e geração técnica de arquivos.
    """
    log_dir = os.path.join(ctx['BASE_DIR'], 'logs')
    os.makedirs(log_dir, exist_ok=True)

    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(1, weight=1)

    header = ttk.Frame(parent)
    header.grid(row=0, column=0, sticky='ew', padx=18, pady=(16, 8))
    ttk.Label(header, text='Modo Desenvolvedor', style='Title.TLabel').pack(side='left')
    ttk.Label(header, text='Diagnóstico, testes e manutenção técnica do sistema', font=('Arial', 10)).pack(side='left', padx=14)

    body = ttk.Frame(parent)
    body.grid(row=1, column=0, sticky='nsew', padx=18, pady=8)
    body.columnconfigure(0, weight=0)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(0, weight=1)

    painel = ttk.LabelFrame(body, text='Ações rápidas')
    painel.grid(row=0, column=0, sticky='ns', padx=(0, 12), pady=0)

    resultado = tk.Text(body, height=28, wrap='word', font=('Consolas', 10))
    resultado.grid(row=0, column=1, sticky='nsew')
    scroll = ttk.Scrollbar(body, orient='vertical', command=resultado.yview)
    scroll.grid(row=0, column=2, sticky='ns')
    resultado.configure(yscrollcommand=scroll.set)

    def append(txt=''):
        resultado.insert('end', txt + '\n')
        resultado.see('end')
        resultado.update_idletasks()

    def limpar():
        resultado.delete('1.0', 'end')

    def verificar_pastas():
        limpar()
        append('VERIFICAÇÃO DE PASTAS')
        append('-'*60)
        pastas = [
            ('Base do projeto', ctx['BASE_DIR']),
            ('Banco de dados', ctx['DATA_DIR']),
            ('PDFs', ctx['PDF_DIR']),
            ('Backups', ctx['BACKUP_DIR']),
            ('Relatórios', ctx['RELATORIO_DIR']),
            ('Modelos de documentos', ctx['MODELOS_DIR']),
            ('Documentos gerados', ctx['DOCS_GERADOS_DIR']),
            ('Assets', ctx['ASSETS_DIR']),
            ('Logs', log_dir),
        ]
        for nome, pasta in pastas:
            exists = os.path.isdir(pasta)
            append(f'{_status_text(exists):<5} {nome:<25} {pasta}')
        _write_log(log_dir, resultado.get('1.0', 'end'))

    def verificar_dependencias():
        limpar()
        append('VERIFICAÇÃO DE DEPENDÊNCIAS')
        append('-'*60)
        deps = [
            ('Python', None),
            ('tkinter', 'tkinter'),
            ('sqlite3', 'sqlite3'),
            ('reportlab', 'reportlab'),
            ('openpyxl', 'openpyxl'),
            ('python-docx', 'docx'),
        ]
        for nome, mod in deps:
            if mod is None:
                append(f'OK    {nome:<15} {sys.version.split()[0]}')
                continue
            try:
                __import__(mod)
                append(f'OK    {nome:<15} instalado')
            except Exception as e:
                append(f'ERRO  {nome:<15} {e}')
        _write_log(log_dir, resultado.get('1.0', 'end'))

    def verificar_banco():
        limpar()
        append('VERIFICAÇÃO DO BANCO SQLITE')
        append('-'*60)
        db = ctx['DB_PATH']
        append(f'Arquivo: {db}')
        if not os.path.exists(db):
            append('ERRO  Banco não encontrado.')
            return
        try:
            con = sqlite3.connect(db)
            cur = con.cursor()
            cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name")
            tabelas = [r[0] for r in cur.fetchall()]
            append(f'OK    Tabelas encontradas: {len(tabelas)}')
            for t in tabelas:
                try:
                    cur.execute(f'SELECT COUNT(*) FROM "{t}"')
                    qtd = cur.fetchone()[0]
                    append(f'      {t:<30} {qtd:>6} registros')
                except Exception:
                    append(f'      {t:<30} não foi possível contar')
            con.close()
        except Exception:
            append('ERRO  Falha ao abrir banco:')
            append(traceback.format_exc())
        _write_log(log_dir, resultado.get('1.0', 'end'))

    def validar_modelos():
        limpar()
        append('VALIDAÇÃO DOS MODELOS WORD')
        append('-'*60)
        pasta = ctx['MODELOS_DIR']
        if not os.path.isdir(pasta):
            append('ERRO  Pasta de modelos não encontrada.')
            return
        arquivos = [a for a in os.listdir(pasta) if a.lower().endswith(('.docx', '.doc'))]
        append(f'Modelos encontrados: {len(arquivos)}')
        for a in sorted(arquivos):
            append(f'OK    {a}')
        _write_log(log_dir, resultado.get('1.0', 'end'))

    def teste_pdf():
        limpar()
        append('TESTE TÉCNICO DE GERAÇÃO DE PDF')
        append('-'*60)
        try:
            from reportlab.lib.pagesizes import A4
            from reportlab.pdfgen import canvas
            path = os.path.join(ctx['PDF_DIR'], 'TESTE_MODO_DESENVOLVEDOR.pdf')
            os.makedirs(ctx['PDF_DIR'], exist_ok=True)
            c = canvas.Canvas(path, pagesize=A4)
            w, h = A4
            c.setFont('Helvetica-Bold', 16)
            c.drawString(50, h-60, 'Sistema RH Izzant - Teste de PDF')
            c.setFont('Helvetica', 10)
            c.drawString(50, h-90, 'Este arquivo foi gerado pelo Modo Desenvolvedor para validar a biblioteca ReportLab.')
            c.drawString(50, h-110, datetime.now().strftime('Data/hora: %d/%m/%Y %H:%M:%S'))
            c.save()
            append(f'OK    PDF gerado: {path}')
        except Exception:
            append('ERRO  Falha ao gerar PDF:')
            append(traceback.format_exc())
        _write_log(log_dir, resultado.get('1.0', 'end'))

    def diagnostico_completo():
        limpar()
        append('DIAGNÓSTICO COMPLETO DO SISTEMA RH IZZANT')
        append('='*72)
        append(datetime.now().strftime('Executado em %d/%m/%Y às %H:%M:%S'))
        append('')

        append('[1] Pastas')
        for nome, pasta in [
            ('Base', ctx['BASE_DIR']), ('Dados', ctx['DATA_DIR']), ('PDFs', ctx['PDF_DIR']),
            ('Backups', ctx['BACKUP_DIR']), ('Modelos', ctx['MODELOS_DIR']), ('Assets', ctx['ASSETS_DIR'])
        ]:
            append(f'{_status_text(os.path.isdir(pasta)):<5} {nome:<10} {pasta}')
        append('')

        append('[2] Banco')
        if os.path.exists(ctx['DB_PATH']):
            try:
                con = sqlite3.connect(ctx['DB_PATH'])
                cur = con.cursor()
                cur.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table'")
                append(f'OK    Banco aberto. Tabelas: {cur.fetchone()[0]}')
                con.close()
            except Exception as e:
                append(f'ERRO  Banco: {e}')
        else:
            append('ERRO  Banco não encontrado')
        append('')

        append('[3] Dependências')
        for nome, mod in [('reportlab','reportlab'), ('openpyxl','openpyxl'), ('python-docx','docx')]:
            try:
                __import__(mod); append(f'OK    {nome}')
            except Exception as e:
                append(f'ERRO  {nome}: {e}')
        append('')

        append('[4] Recursos')
        append(f'{_status_text(os.path.exists(ctx["LOGO_APP"])):<5} Logo da aplicação')
        modelos = [a for a in os.listdir(ctx['MODELOS_DIR']) if a.lower().endswith('.docx')] if os.path.isdir(ctx['MODELOS_DIR']) else []
        append(f'OK    Modelos Word: {len(modelos)}')
        append('')
        append('Status geral: diagnóstico concluído. Verifique linhas com ERRO, se houver.')
        _write_log(log_dir, resultado.get('1.0', 'end'))

    def abrir_pasta(path):
        os.makedirs(path, exist_ok=True)
        webbrowser.open(path)

    def abrir_log():
        abrir_pasta(log_dir)

    botoes = [
        ('Diagnóstico completo', diagnostico_completo),
        ('Verificar banco SQLite', verificar_banco),
        ('Verificar pastas', verificar_pastas),
        ('Verificar dependências', verificar_dependencias),
        ('Validar modelos Word', validar_modelos),
        ('Testar geração PDF', teste_pdf),
        ('Abrir pasta do projeto', lambda: abrir_pasta(ctx['BASE_DIR'])),
        ('Abrir pasta PDFs', lambda: abrir_pasta(ctx['PDF_DIR'])),
        ('Abrir pasta Logs', abrir_log),
        ('Limpar resultado', limpar),
    ]
    for texto, cmd in botoes:
        ttk.Button(painel, text=texto, command=cmd, width=26).pack(fill='x', padx=10, pady=5)

    ttk.Label(parent, text='Este módulo é técnico e não altera dados de produção, exceto quando gerar arquivos de teste em PDFs/logs.', font=('Arial', 9)).grid(row=2, column=0, sticky='w', padx=18, pady=(0,12))

    diagnostico_completo()
